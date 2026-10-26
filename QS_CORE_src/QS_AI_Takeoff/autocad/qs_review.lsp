;;; QS_REVIEW - navigate flagged records from a review file and update their status.
;;; Review file (written by the app): one record per line  <record_id>|<handles comma-sep>|<reason>
;;; Status decisions are written to <file>_decisions.txt as  <record_id>|<APPROVED|REJECTED|NEEDS_REVIEW>
;;; Source entities are only highlighted (zoom + temporary selection), never modified.
(vl-load-com)

(defun qsr:split (s d / out pos)
  (setq out '())
  (while (setq pos (vl-string-search d s))
    (setq out (cons (substr s 1 pos) out) s (substr s (+ pos 1 (strlen d)))))
  (reverse (cons s out)))

(defun qsr:zoom-handles (handles / ents e ss minp maxp obj ll ur)
  (setq ss (ssadd))
  (foreach h handles
    (if (setq e (handent h)) (ssadd e ss)))
  (if (> (sslength ss) 0)
    (progn
      (setq minp nil maxp nil)
      (foreach h handles
        (if (setq e (handent h))
          (progn (setq obj (vlax-ename->vla-object e))
                 (vla-getboundingbox obj 'll 'ur)
                 (setq ll (vlax-safearray->list ll) ur (vlax-safearray->list ur))
                 (setq minp (if minp (mapcar 'min minp ll) ll) maxp (if maxp (mapcar 'max maxp ur) ur)))))
      (command "_.ZOOM" "_W" (mapcar '- minp '(500 500 0)) (mapcar '+ maxp '(500 500 0)))
      (sssetfirst nil ss)))
  (sslength ss))

(defun c:QS_REVIEW (/ path f out line rec handles ans n olderr)
  (setq olderr *error*)
  (defun *error* (m) (if f (close f)) (if out (close out)) (sssetfirst nil nil) (setq *error* olderr)
    (princ (strcat "\nQS_REVIEW: " m)) (princ))
  (setq path (getfiled "Select QS review file" (getvar "DWGPREFIX") "txt" 4))
  (if (null path) (exit))
  (setq f (open path "r") out (open (strcat (vl-filename-directory path) "\\" (vl-filename-base path) "_decisions.txt") "a") n 0)
  (while (setq line (read-line f))
    (setq rec (qsr:split line "|") handles (qsr:split (nth 1 rec) ","))
    (if (> (qsr:zoom-handles handles) 0)
      (progn
        (prompt (strcat "\n[" (car rec) "] " (if (> (length rec) 2) (nth 2 rec) "")))
        (initget "Approve Reject Skip Quit")
        (setq ans (getkword "\nDecision [Approve/Reject/Skip/Quit] <Skip>: "))
        (cond ((= ans "Approve") (write-line (strcat (car rec) "|APPROVED") out) (setq n (1+ n)))
              ((= ans "Reject") (write-line (strcat (car rec) "|REJECTED") out) (setq n (1+ n)))
              ((= ans "Quit") (setq line nil)))
        (if (= ans "Quit") (progn (close f) (setq f nil) (close out) (setq out nil) (sssetfirst nil nil)
                                   (prompt (strcat "\nQS_REVIEW: " (itoa n) " decisions saved.")) (exit))))
      (prompt (strcat "\n[" (car rec) "] source handles not found in this drawing - skipped."))))
  (close f) (setq f nil) (close out) (setq out nil)
  (sssetfirst nil nil)
  (setq *error* olderr)
  (prompt (strcat "\nQS_REVIEW: " (itoa n) " decisions saved. Import them in the app (Review > Import decisions)."))
  (princ))

;;; QS_GOTO - zoom to one or more handles typed at the prompt (comma separated).
(defun c:QS_GOTO (/ s)
  (setq s (getstring "\nHandle(s), comma separated: "))
  (if (= (qsr:zoom-handles (qsr:split s ",")) 0) (prompt "\nHandle not found."))
  (princ))
(princ "\nQS_REVIEW / QS_GOTO loaded.")
(princ)
