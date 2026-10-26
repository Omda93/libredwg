;;; QS_DRAW - draw application-generated geometry from a script file written by the Python app.
;;; The script is a plain text file, one record per line:
;;;   LAYER|<name>|<color>
;;;   LINE|<layer>|<x1>|<y1>|<x2>|<y2>|<id>
;;;   PLINE|<layer>|<id>|x1,y1;x2,y2;...     (closed)
;;;   TEXT|<layer>|<x>|<y>|<height>|<text>|<id>
;;; Every created entity gets XData (app QS_TAKEOFF) with its id so QS_PURGE can remove / update it.
;;; Idempotent: before drawing, existing entities with the same id are erased.
;;; Source entities are never touched.
(vl-load-com)

(defun qs:split (s d / out pos)
  (setq out '())
  (while (setq pos (vl-string-search d s))
    (setq out (cons (substr s 1 pos) out) s (substr s (+ pos 1 (strlen d)))))
  (reverse (cons s out)))

(defun qs:regapp () (if (not (tblsearch "APPID" "QS_TAKEOFF")) (regapp "QS_TAKEOFF")))

(defun qs:xd (id) (list (list -3 (list "QS_TAKEOFF" (cons 1000 id)))))

(defun qs:layer (name color / l)
  (if (not (tblsearch "LAYER" name))
    (entmake (list '(0 . "LAYER") '(100 . "AcDbSymbolTableRecord") '(100 . "AcDbLayerTableRecord")
                   (cons 2 name) '(70 . 0) (cons 62 (atoi color)) '(6 . "Continuous")))))

;;; ssget's XData filter matches the APP NAME only, not the stored value, so the id -> entity
;;; map is built once per run by reading each QS entity's XData string.
(setq *qs-idmap* nil)

(defun qs:xd-id (ent / xd)
  (setq xd (cdr (assoc -3 (entget ent '("QS_TAKEOFF")))))
  (if xd (cdr (assoc 1000 (cdr (car xd))))))

(defun qs:build-idmap (/ ss i e id)
  (setq *qs-idmap* nil ss (ssget "_X" '((-3 ("QS_TAKEOFF")))) i 0)
  (if ss
    (repeat (sslength ss)
      (setq e (ssname ss i) id (qs:xd-id e))
      (if id (setq *qs-idmap* (cons (cons id e) *qs-idmap*)))
      (setq i (1+ i)))))

(defun qs:erase-id (id / pair)
  (while (setq pair (assoc id *qs-idmap*))
    (if (entget (cdr pair)) (entdel (cdr pair)))
    (setq *qs-idmap* (vl-remove pair *qs-idmap*))))

(defun qs:draw-line (f)
  (qs:erase-id (nth 6 f))
  (entmake (append (list '(0 . "LINE") (cons 8 (nth 1 f))
                         (list 10 (atof (nth 2 f)) (atof (nth 3 f)) 0.0)
                         (list 11 (atof (nth 4 f)) (atof (nth 5 f)) 0.0))
                   (qs:xd (nth 6 f)))))

(defun qs:draw-pline (f / pts p)
  (qs:erase-id (nth 2 f))
  (setq pts (mapcar '(lambda (s) (setq p (qs:split s ",")) (list 10 (atof (car p)) (atof (cadr p))))
                    (qs:split (nth 3 f) ";")))
  (entmake (append (list '(0 . "LWPOLYLINE") '(100 . "AcDbEntity") (cons 8 (nth 1 f))
                         '(100 . "AcDbPolyline") (cons 90 (length pts)) '(70 . 1))
                   pts (qs:xd (nth 2 f)))))

(defun qs:draw-text (f)
  (qs:erase-id (nth 6 f))
  (entmake (append (list '(0 . "TEXT") (cons 8 (nth 1 f)) (list 10 (atof (nth 2 f)) (atof (nth 3 f)) 0.0)
                         (cons 40 (atof (nth 4 f))) (cons 1 (nth 5 f)))
                   (qs:xd (nth 6 f)))))

;;; Non-interactive core: draw every record of a script file. Returns the number of entities made.
;;; Existing entities carrying the same id are erased first (idempotent re-run).
(defun qs:draw-file (path / f line fields n)
  (qs:regapp)
  (qs:build-idmap)
  (setq n 0)
  (if (setq f (open path "r"))
    (progn
      (while (setq line (read-line f))
        (setq fields (qs:split line "|"))
        (cond ((= (car fields) "LAYER") (qs:layer (nth 1 fields) (nth 2 fields)))
              ((= (car fields) "LINE") (if (qs:draw-line fields) (setq n (1+ n))))
              ((= (car fields) "PLINE") (if (qs:draw-pline fields) (setq n (1+ n))))
              ((= (car fields) "TEXT") (if (qs:draw-text fields) (setq n (1+ n))))))
      (close f))
    (prompt (strcat "\nQS_DRAW: cannot open " path)))
  n)

;;; Non-interactive purge of application-generated entities only. Returns count erased.
(defun qs:purge-all (/ ss i)
  (qs:regapp)
  (setq ss (ssget "_X" '((-3 ("QS_TAKEOFF")))) i 0)
  (if ss (repeat (sslength ss) (entdel (ssname ss i)) (setq i (1+ i))))
  i)

(defun c:QS_DRAW (/ path n olderr oldcmd)
  (setq olderr *error* oldcmd (getvar "CMDECHO"))
  (defun *error* (m) (setvar "CMDECHO" oldcmd) (setq *error* olderr)
    (princ (strcat "\nQS_DRAW error: " m)) (princ))
  (setvar "CMDECHO" 0)
  (setq path (getfiled "Select QS draw script" (getvar "DWGPREFIX") "txt" 4))
  (if path
    (progn (setq n (qs:draw-file path))
           (prompt (strcat "\nQS_DRAW: " (itoa n) " generated entities created/updated."))))
  (setvar "CMDECHO" oldcmd) (setq *error* olderr)
  (princ))

;;; QS_PURGE - remove ONLY application-generated entities (those carrying QS_TAKEOFF XData).
(defun c:QS_PURGE (/ ss ans)
  (qs:regapp)
  (setq ss (ssget "_X" '((-3 ("QS_TAKEOFF")))))
  (if (null ss) (prompt "\nNo QS generated entities found.")
    (progn
      (initget "Yes No")
      (setq ans (getkword (strcat "\nErase " (itoa (sslength ss)) " QS generated entities? [Yes/No] <No>: ")))
      (if (= ans "Yes")
        (prompt (strcat "\n" (itoa (qs:purge-all)) " QS generated entities erased."))
        (prompt "\nCancelled."))))
  (princ))
(princ "\nQS_DRAW / QS_PURGE loaded.")
(princ)
