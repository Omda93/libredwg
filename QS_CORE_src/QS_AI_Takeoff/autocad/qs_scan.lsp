;;; QS_SCAN - export LINE / LWPOLYLINE / ARC / CIRCLE geometry and INSERT blocks (with attributes) to JSON.
;;; Read-only: no entity is modified. Output: <dwg folder>\<dwgname>_qs_scan.json
;;; TEXT / MTEXT (and room-tag attributes) are exported too - they give the room names.
;;; Usage: APPLOAD this file, then type QS_SCAN, select entities (or press Enter for ALL).
;;; The Python app calls (qs:scan-file "<json>" "<layer patterns>") directly - it runs INSIDE AutoCAD,
;;; which is far faster than reading entities one by one over COM.
;;; Curves: arcs, circles and bulged polyline segments are written as short pieces flagged "arc":true
;;; (room outlines); a bulged segment also keeps its straight chord flagged "chord":true (wall pairing).
(vl-load-com)

(defun qs:esc (s / cs)
  ;; escape backslash, double quote and control characters for JSON - linear, and a no-op for plain strings
  ;; (the former char-by-char strcat loop cost ~40 s on 10,000 MTEXTs)
  (setq cs (vl-string->list s))
  (if (vl-some '(lambda (c) (or (< c 32) (= c 34) (= c 92))) cs)
    (vl-list->string
      (apply 'append
             (mapcar '(lambda (c) (cond ((= c 92) '(92 92)) ((= c 34) '(92 34)) ((< c 32) '(32)) (t (list c))))
                     cs)))
    s))

(defun qs:want-layer (lay)
  ;; layer filter of the current scan (*qs-pats*, nil = all). Layer "0" inside a block takes the
  ;; INSERT's layer, so it is always kept there.
  (or (null *qs-pats*) (= lay "0") (wcmatch (strcase lay) *qs-pats*)))

(defun qs:esc-layer (s / hit)
  ;; layer names repeat thousands of times - escape each one once
  (if (setq hit (assoc s *qs-lc*))
    (cdr hit)
    (progn (setq hit (qs:esc s) *qs-lc* (cons (cons s hit) *qs-lc*)) hit)))

(defun qs:num (x / s)
  ;; always valid JSON, whatever DIMZIN says (".5" -> "0.5", "-.5" -> "-0.5")
  (setq s (rtos x 2 6))
  (cond ((= (substr s 1 1) ".") (strcat "0" s))
        ((= (substr s 1 2) "-.") (strcat "-0" (substr s 2)))
        (t s)))

(defun qs:line-json (h lay p1 p2 src)
  (qs:line-json2 h lay p1 p2 src ""))

(defun qs:line-json2 (h lay p1 p2 src extra)
  (strcat "{\"handle\":\"" h "\",\"layer\":\"" (qs:esc-layer lay) "\",\"x1\":" (qs:num (car p1))
          ",\"y1\":" (qs:num (cadr p1)) ",\"x2\":" (qs:num (car p2)) ",\"y2\":" (qs:num (cadr p2))
          ",\"src\":\"" src "\"" extra "}"))

(defun qs:out (txt f)
  (write-line (strcat (if *qs-first* "" ",") txt) f)
  (setq *qs-first* nil))

(defun qs:attrs (ent / out e ed)
  ;; ATTRIB sub-entities that follow the INSERT (pure DXF: also works in accoreconsole, no COM)
  (setq out "" ed (entget ent))
  (if (= 1 (cdr (assoc 66 ed)))
    (progn
      (setq e (entnext ent))
      (while (and e (= "ATTRIB" (cdr (assoc 0 (setq ed (entget e))))))
        (setq out (strcat out (if (= out "") "" ",") "\"" (qs:esc (cdr (assoc 2 ed))) "\":\""
                          (qs:esc (cdr (assoc 1 ed))) "\"")
              e (entnext e)))))
  (strcat "{" out "}"))

(defun qs:eff-name (bname / hit br x h)
  ;; effective name of a dynamic block (*U copy -> original) from the block record's AcDbBlockRepBTag
  (if (setq hit (assoc bname *qs-eff*))
    (cdr hit)
    (progn
      (setq h bname)
      (if (= (substr bname 1 1) "*")
        (progn
          (setq x (vl-catch-all-apply
                    '(lambda ()
                       (setq br (cdr (assoc 330 (entget (tblobjname "BLOCK" bname)))))
                       (setq x (cdr (assoc -3 (entget br '("AcDbBlockRepBTag")))))
                       (if (and x (setq x (cdr (assoc 1005 (cdar x)))) (setq x (handent x)))
                         (cdr (assoc 2 (entget x)))))))
          (if (and x (not (vl-catch-all-error-p x))) (setq h x))))
      (setq *qs-eff* (cons (cons bname h) *qs-eff*))
      h)))

(defun qs:normal (ed) (if (assoc 210 ed) (cdr (assoc 210 ed)) '(0.0 0.0 1.0)))

(defun qs:nseg (sweep) (max 2 (fix (+ 0.999999 (/ (abs sweep) (/ pi 36.0))))))   ; 5 degree pieces

(defun qs:arc-pts (c r a0 sweep / n k pts)
  ;; OCS points of an arc: centre c, radius r, start a0, signed sweep
  (setq n (qs:nseg sweep) k 0 pts nil)
  (repeat (1+ n)
    (setq pts (cons (polar c (+ a0 (* sweep (/ (float k) n))) r) pts) k (1+ k)))
  (reverse pts))

(defun qs:bulge-pts (p1 p2 b / ch th r mx my d ux uy s cx cy a0 pts)
  ;; OCS points of a bulged polyline segment p1 -> p2 (same maths as engine/curves.py)
  (setq ch (distance (list (car p1) (cadr p1)) (list (car p2) (cadr p2))))
  (if (or (< (abs b) 1e-9) (< ch 1e-9))
    (list p1 p2)
    (progn
      (setq th (* 4.0 (atan b)) r (/ ch (* 2.0 (sin (/ (abs th) 2.0))))
            mx (/ (+ (car p1) (car p2)) 2.0) my (/ (+ (cadr p1) (cadr p2)) 2.0)
            d (sqrt (max (- (* r r) (* (/ ch 2.0) (/ ch 2.0))) 0.0))
            ux (/ (- (car p2) (car p1)) ch) uy (/ (- (cadr p2) (cadr p1)) ch)
            s (if (eq (> b 0) (< (abs th) pi)) 1.0 -1.0)
            cx (- mx (* s uy d)) cy (+ my (* s ux d))
            a0 (atan (- (cadr p1) cy) (- (car p1) cx)))
      (setq pts (qs:arc-pts (list cx cy (caddr p1)) r a0 th))
      (append (list p1) (cdr (reverse (cdr (reverse pts)))) (list p2)))))

(defun qs:write-pieces (pts nrm h lay src f / k a b)
  ;; consecutive OCS points -> WCS pieces flagged "arc":true
  (setq k 0)
  (while (cdr pts)
    (setq a (trans (car pts) nrm 0) b (trans (cadr pts) nrm 0))
    (if (> (distance a b) 1e-9)
      (qs:out (qs:line-json2 (strcat h "~" (itoa k)) lay a b src ",\"arc\":true") f))
    (setq pts (cdr pts) k (1+ k))))

(defun qs:write-pline (ent f / ed nrm elev verts i closed v1 v2 p1 p2 h lay b rest)
  ;; vertices are OCS -> converted with the entity normal (handles mirrored polylines); bulges kept
  (setq ed (entget ent) h (cdr (assoc 5 ed)) lay (cdr (assoc 8 ed)) nrm (qs:normal ed)
        elev (if (assoc 38 ed) (cdr (assoc 38 ed)) 0.0) verts nil)
  (foreach x ed
    (cond ((= (car x) 10) (setq verts (cons (list (cadr x) (caddr x) elev 0.0) verts)))
          ((and (= (car x) 42) verts)
           (setq verts (cons (list (car (car verts)) (cadr (car verts)) elev (cdr x)) (cdr verts))))))
  (setq verts (reverse verts))
  ;; linear walk (the former (nth i verts) loop was quadratic: minutes on detail polylines with many vertices)
  (setq closed (= 1 (logand 1 (cdr (assoc 70 ed)))) i 0 rest verts)
  (while rest
    (setq v1 (car rest) v2 (cond ((cdr rest) (cadr rest)) (closed (car verts)) (t nil)))
    (if v2
      (progn
        (setq b (cadddr v1) p1 (list (car v1) (cadr v1) elev) p2 (list (car v2) (cadr v2) elev))
        (if (< (abs b) 1e-9)
          (qs:out (qs:line-json (strcat h "." (itoa i)) lay (trans p1 nrm 0) (trans p2 nrm 0) "LWPOLYLINE") f)
          (progn
            (qs:out (qs:line-json2 (strcat h "." (itoa i)) lay (trans p1 nrm 0) (trans p2 nrm 0) "LWPOLYLINE"
                                   ",\"chord\":true") f)
            (qs:write-pieces (qs:bulge-pts p1 p2 b) nrm (strcat h "." (itoa i)) lay "LWPOLYLINE" f)))))
    (setq rest (cdr rest) i (1+ i))))

(defun qs:write-geom (ent f / ed typ nrm sw)
  ;; LINE / LWPOLYLINE / ARC / CIRCLE -> line records. Returns T when something was written.
  (setq ed (entget ent) typ (cdr (assoc 0 ed)))
  (cond ((= typ "LINE")
         (qs:out (qs:line-json (cdr (assoc 5 ed)) (cdr (assoc 8 ed)) (cdr (assoc 10 ed)) (cdr (assoc 11 ed)) "LINE") f)
         t)
        ((= typ "LWPOLYLINE") (qs:write-pline ent f) t)
        ((= typ "ARC")
         (setq nrm (qs:normal ed) sw (- (cdr (assoc 51 ed)) (cdr (assoc 50 ed))))
         (while (<= sw 0.0) (setq sw (+ sw (* 2.0 pi))))
         (qs:write-pieces (qs:arc-pts (cdr (assoc 10 ed)) (cdr (assoc 40 ed)) (cdr (assoc 50 ed)) sw)
                          nrm (cdr (assoc 5 ed)) (cdr (assoc 8 ed)) "ARC" f)
         t)
        ((= typ "CIRCLE")
         (qs:write-pieces (qs:arc-pts (cdr (assoc 10 ed)) (cdr (assoc 40 ed)) 0.0 (* 2.0 pi))
                          (qs:normal ed) (cdr (assoc 5 ed)) (cdr (assoc 8 ed)) "CIRCLE" f)
         t)))

(defun qs:text-json (ent / ed typ p s x)
  ;; TEXT (insert in OCS) / MTEXT (insert in WCS, text in groups 3 + 1) -> {"handle","layer","x","y","h","text"}
  (setq ed (entget ent) typ (cdr (assoc 0 ed)))
  (if (= typ "MTEXT")
    (progn
      (setq s "" p (cdr (assoc 10 ed)))
      (foreach x ed (if (= (car x) 3) (setq s (strcat s (cdr x)))))
      (setq s (strcat s (cdr (assoc 1 ed)))))
    (setq s (cdr (assoc 1 ed)) p (trans (cdr (assoc 10 ed)) (qs:normal ed) 0)))
  (strcat "{\"handle\":\"" (cdr (assoc 5 ed)) "\",\"layer\":\"" (qs:esc-layer (cdr (assoc 8 ed)))
          "\",\"x\":" (qs:num (car p)) ",\"y\":" (qs:num (cadr p))
          ",\"h\":" (qs:num (if (assoc 40 ed) (cdr (assoc 40 ed)) 0.0))
          ",\"text\":\"" (qs:esc s) "\",\"mtext\":" (if (= typ "MTEXT") "true" "false") "}"))

(defun qs:is-text (ent) (member (cdr (assoc 0 (entget ent))) '("TEXT" "MTEXT")))

(defun qs:insert-json (ent / ed p nrm bname ename)
  ;; Same fields as the COM reader: block = definition name (anonymous *U for dynamic blocks),
  ;; x/y in the owner's WCS, rot radians, nz = -1 for mirrored (OCS 0,0,-1) inserts.
  (setq ed (entget ent) nrm (qs:normal ed)
        p (trans (cdr (assoc 10 ed)) nrm 0) bname (cdr (assoc 2 ed)))
  (setq ename (qs:eff-name bname))
  (if (not (member bname *qs-names*)) (setq *qs-names* (cons bname *qs-names*)))
  (strcat "{\"handle\":\"" (cdr (assoc 5 ed)) "\",\"name\":\"" (qs:esc ename) "\",\"block\":\"" (qs:esc bname)
          "\",\"layer\":\"" (qs:esc-layer (cdr (assoc 8 ed))) "\",\"x\":" (qs:num (car p)) ",\"y\":" (qs:num (cadr p))
          ",\"sx\":" (qs:num (if (assoc 41 ed) (cdr (assoc 41 ed)) 1.0))
          ",\"sy\":" (qs:num (if (assoc 42 ed) (cdr (assoc 42 ed)) 1.0))
          ",\"rot\":" (qs:num (if (assoc 50 ed) (cdr (assoc 50 ed)) 0.0))
          ",\"rotation\":" (qs:num (if (assoc 50 ed) (cdr (assoc 50 ed)) 0.0))
          ",\"nz\":" (if (< (caddr nrm) 0) "-1.0" "1.0")
          ",\"attrs\":" (qs:attrs ent) "}"))

(defun qs:write-defs (f prog / done name bd e first-def base xref n ed typ skipped ins txt fam wrote famskip)
  ;; Write every referenced block definition (recursively, read-only) as "block_defs":{...}
  ;; Geometry inside definitions obeys the same layer filter as model space: door / window families
  ;; carry 10,000+ detail lines each, which used to make this stage take tens of minutes.
  (setq done '() first-def t n 0 skipped 0 famskip 0 *qs-empty* nil)
  (write-line "\"block_defs\":{" f)
  (while *qs-names*
    (setq name (car *qs-names*) *qs-names* (cdr *qs-names*))
    (if (and (not (member name done)) (setq bd (tblsearch "BLOCK" name)))
      (progn
        (setq done (cons name done) base (cdr (assoc 10 bd)) xref (/= 0 (logand 4 (cdr (assoc 70 bd)))) n (1+ n))
        (qs:stage prog "defs" n (+ n (length *qs-names*)))
        (write-line (strcat (if first-def "" ",") "\"" (qs:esc name) "\":{\"base\":[" (qs:num (car base)) ","
                            (qs:num (cadr base)) "],\"xref\":" (if xref "true" "false") ",\"lines\":[") f)
        (setq first-def nil *qs-first* t ins nil txt nil fam (qs:eff-name name) wrote 0)
        ;; Dynamic blocks: every door / window instance is its own anonymous *U copy (thousands, 10,000+
        ;; detail lines each). With a layer filter, once one copy of a family gave no wanted geometry
        ;; and no nested blocks, its other copies are not walked again.
        (if (and *qs-pats* (member fam *qs-empty*))
          (setq xref t famskip (1+ famskip)))
        (if (not xref)                                   ; ONE walk per definition (was three)
          (progn
            (setq e (cdr (assoc -2 bd)))
            (while e
              (setq ed (entget e) typ (cdr (assoc 0 ed)))
              (cond ((= typ "INSERT") (setq ins (cons e ins)))
                    ((member typ '("TEXT" "MTEXT")) (setq txt (cons e txt)))
                    ((member typ '("LINE" "LWPOLYLINE" "ARC" "CIRCLE"))
                     (if (qs:want-layer (cdr (assoc 8 ed)))
                       (progn (qs:write-geom e f) (setq wrote (1+ wrote)))
                       (setq skipped (1+ skipped)))))
              (setq e (entnext e)))
            (if (and *qs-pats* (= wrote 0) (null ins) (not (member fam *qs-empty*)))
              (setq *qs-empty* (cons fam *qs-empty*)))))
        (write-line "],\"inserts\":[" f)
        (setq *qs-first* t)
        (foreach e (reverse ins) (qs:out (qs:insert-json e) f))
        (write-line "],\"texts\":[" f)
        (setq *qs-first* t)
        (foreach e (reverse txt) (qs:out (qs:text-json e) f))
        (write-line "]}" f))))
  (write-line "}," f)
  (write-line (strcat "\"defs_skipped_entities\":" (itoa skipped) ",\"defs_family_skipped\":" (itoa famskip)) f)
  n)

(defun qs:progress (path i n / pf)
  (qs:stage path "lines" i n))

(defun qs:stage (path stage i n / pf)
  ;; progress file read by the app: "<stage>:<i>/<n>"  (stages: start, lines, blocks, texts, defs)
  (if (and path (setq pf (open path "w")))
    (progn (write-line (strcat stage ":" (itoa i) "/" (itoa n)) pf) (close pf))))

(defun qs:write-scan (ss path prog / f i ent ed typ lines blocks defs n)
  ;; the JSON body shared by QS_SCAN (interactive), qs:scan-file / qs:scan-window (app) and the
  ;; background accoreconsole scan. ONE pass over the selection (was three).
  (setq f (open path "w") lines 0 blocks 0 i 0 n (sslength ss) *qs-lc* nil *qs-eff* nil
        *qs-blk* nil *qs-txt* nil)
  (write-line (strcat "{\"meta\":{\"source\":\"" (qs:esc (getvar "DWGNAME")) "\",\"insunits\":"
                      (itoa (getvar "INSUNITS")) ",\"acadver\":\"" (getvar "ACADVER")
                      "\",\"reader\":\"qs_scan.lsp\",\"selected\":" (itoa n) "}," ) f)
  (write-line "\"lines\":[" f)
  (setq *qs-first* t *qs-names* '())
  (repeat n
    (setq ent (ssname ss i) ed (entget ent) typ (cdr (assoc 0 ed)))
    (cond ((= typ "INSERT") (setq *qs-blk* (cons ent *qs-blk*)))
          ((member typ '("TEXT" "MTEXT")) (setq *qs-txt* (cons ent *qs-txt*)))
          ((qs:write-geom ent f) (setq lines (1+ lines))))
    (setq i (1+ i))
    (if (= 0 (rem i 2000)) (qs:progress prog i n)))
  (write-line "]," f)
  (write-line "\"blocks\":[" f)
  (setq *qs-first* t i 0)
  (foreach ent (reverse *qs-blk*)
    (qs:out (qs:insert-json ent) f)
    (setq blocks (1+ blocks) i (1+ i))
    (if (= 0 (rem i 500)) (qs:stage prog "blocks" i (length *qs-blk*))))
  (write-line "]," f)
  (write-line "\"texts\":[" f)
  (setq *qs-first* t i 0)
  (foreach ent (reverse *qs-txt*)
    (qs:out (qs:text-json ent) f)
    (setq i (1+ i))
    (if (= 0 (rem i 1000)) (qs:stage prog "texts" i (length *qs-txt*))))
  (write-line "]," f)
  (qs:stage prog "defs" 0 (length *qs-names*))
  (setq defs (qs:write-defs f prog))
  (write-line "}" f)
  (close f)
  (setq *qs-blk* nil *qs-txt* nil)
  (list lines blocks defs))

(defun qs:pats (pats)
  ;; "A,B,*COL*" -> upper-case wcmatch pattern for the definition filter (nil = all layers)
  (if (and pats (/= pats "")) (strcase pats) nil))

(defun qs:filter (pats)
  ;; ssget filter: geometry on the chosen layers + every INSERT / TEXT / MTEXT (walls inside blocks, names)
  (if (and pats (/= pats ""))
    (list '(-4 . "<OR") '(0 . "INSERT,TEXT,MTEXT")
          '(-4 . "<AND") '(0 . "LINE,LWPOLYLINE,ARC,CIRCLE") (cons 8 pats) '(-4 . "AND>")
          '(-4 . "OR>"))
    '((0 . "LINE,LWPOLYLINE,ARC,CIRCLE,INSERT,TEXT,MTEXT"))))

(defun qs:finish (path res / df)
  (setq df (open (strcat path ".done") "w"))
  (write-line (if res
                (strcat "OK " (itoa (car res)) " " (itoa (cadr res)) " " (itoa (caddr res)))
                "EMPTY") df)
  (close df))

(defun qs:scan-file (path pats / ss res olderr)
  ;; Non-interactive scan of MODEL SPACE for the app. pats = layer wildcard list ("" = all layers);
  ;; INSERTs are always read (walls inside blocks). Writes <path>.done = "OK lines blocks defs" or "ERROR msg".
  ;; Used inside AutoCAD (SendCommand) and by the background accoreconsole scan (no UI).
  (setq olderr *error*)
  (defun *error* (m / ef)
    (if (setq ef (open (strcat path ".done") "w")) (progn (write-line (strcat "ERROR " m) ef) (close ef)))
    (setq *error* olderr) (princ))
  (qs:stage (strcat path ".prog") "start" 0 0)          ; tells the app the routine is running
  (setq *qs-pats* (qs:pats pats))
  (setq ss (ssget "_X" (cons '(410 . "Model") (qs:filter pats))))
  (setq res (if ss (qs:write-scan ss path (strcat path ".prog"))))
  (qs:finish path res)
  (setq *error* olderr *qs-pats* nil)
  (princ))


(defun qs:scan-window (path pats / ss res olderr)
  ;; The user drags a window (or picks objects) in AutoCAD; only that part of the plan is scanned.
  ;; Writes <path>.window = "x1 y1 x2 y2" (box of the selection) for the app's preview.
  (setq olderr *error*)
  (defun *error* (m / ef)
    (if (setq ef (open (strcat path ".done") "w"))
      (progn (write-line (if (wcmatch (strcase m) "*CANCEL*,*QUIT*") "CANCEL" (strcat "ERROR " m)) ef) (close ef)))
    (setq *error* olderr *qs-pats* nil) (princ))
  (qs:stage (strcat path ".prog") "select" 0 0)
  (setq *qs-pats* (qs:pats pats))
  (prompt "\nQS CORE - select the area to take off: drag a window (left->right) or crossing (right->left), Enter when done: ")
  (setq ss (ssget (qs:filter pats)))
  (if ss
    (progn
      (qs:box-file ss (strcat path ".window"))
      (qs:stage (strcat path ".prog") "start" 0 0)
      (setq res (qs:write-scan ss path (strcat path ".prog"))))
    (setq res nil))
  (qs:finish path res)
  (setq *error* olderr *qs-pats* nil)
  (princ))

(defun qs:box-file (ss out / i e mn mx bf a)
  ;; extents of the selection from the DXF points 10 / 11 (no COM)
  (setq i 0 mn nil mx nil)
  (repeat (sslength ss)
    (setq e (entget (ssname ss i)) i (1+ i))
    (foreach g e
      (if (member (car g) '(10 11))
        (progn
          (setq a (cdr g))
          (setq mn (if mn (list (min (car mn) (car a)) (min (cadr mn) (cadr a))) (list (car a) (cadr a)))
                mx (if mx (list (max (car mx) (car a)) (max (cadr mx) (cadr a))) (list (car a) (cadr a))))))))
  (if (and mn (setq bf (open out "w")))
    (progn (write-line (strcat (qs:num (car mn)) " " (qs:num (cadr mn)) " " (qs:num (car mx)) " "
                               (qs:num (cadr mx))) bf)
           (close bf))))

(defun c:QS_SCAN (/ ss path res dwg olderr)
  (setq olderr *error*)
  (defun *error* (m) (setq *error* olderr *qs-pats* nil) (princ (strcat "\nQS_SCAN error: " m)) (princ))
  (prompt "\nSelect entities to scan (window / crossing, Enter = whole drawing): ")
  (setq *qs-pats* nil)
  (setq ss (ssget '((0 . "LINE,LWPOLYLINE,ARC,CIRCLE,INSERT,TEXT,MTEXT"))))
  (if (null ss) (setq ss (ssget "_X" '((410 . "Model") (0 . "LINE,LWPOLYLINE,ARC,CIRCLE,INSERT,TEXT,MTEXT")))))
  (if (null ss) (progn (prompt "\nNothing to scan.") (exit)))
  (setq dwg (vl-filename-base (getvar "DWGNAME")))
  (setq path (strcat (getvar "DWGPREFIX") dwg "_qs_scan.json"))
  (setq res (qs:write-scan ss path nil))
  (setq *error* olderr)
  (prompt (strcat "\nQS_SCAN: " (itoa (car res)) " line entities, " (itoa (cadr res)) " blocks, "
                  (itoa (caddr res)) " block definitions -> " path))
  (prompt (strcat "\nDrawing INSUNITS = " (itoa (getvar "INSUNITS")) " (0 = unitless: confirm units in the app)."))
  (princ))
(princ "\nQS_SCAN loaded. Type QS_SCAN to export geometry.")
(princ)

