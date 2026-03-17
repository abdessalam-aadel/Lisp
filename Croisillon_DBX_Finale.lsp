;;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Croisillon_DBX_Tonkpi : Small lisp to Creat quadrillage 
;;; + modify LineType + Modify the length of Riverain segments 
;;; + Replace Mtext using Regex ... in Multiple DWG file
;;; Author: Abdessalam Aadel © 2026
;;; GitHub: https://github.com/abdessalam-aadel/Lisp
;;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;;  Start function Create-Croix
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun Create-Croix (ms ech pl
					/ 	minPt maxPt step size h
						p1 p2 p3 p4 pp1 pp2 pp3 pp4
						cp1 cp2
						ang1 ang2 ang3 ang4
						x y xx yy p pt1 pt2 ptt1
						x0 y0 p0 dist newPt pm 
					)

	(if (null pl)
		(progn
		  (princ "\nAucune polyligne fermée détectée.")
		  (exit)
		)
	)
	
	;; Start Get Bounding Box of polyligne
	(vla-getBoundingBox pl 'minPt 'maxPt)

	(setq p2 (SafeArray->List minPt))		;p2 = (xmin ymin zmin) → bottom-left corner
	(setq pp4 (SafeArray->List maxPt))		;pp4 = (xmax ymax zmax) → top-right corner
	(setq p1 (list (car pp4) (cadr p2)))	;p1 = (xmax, ymin) → bottom-right corner
	(setq p3 p2)							;p3 = p2
	(setq p4 (list (car p2) (cadr pp4)))	;p4 = (xmin, ymax) → top-left corner
	
	(setq step (* ech 0.10))
	(setq size (* ech 0.0025))
	(setq h (* 0.002 ech))
	
	;Checks whether points are equal within a tolerance of 0.01 units.
	(if (equal p1 p3 0.01)
		(setq p3 p4)
	)
	(if (equal p2 p3 0.01)
		(progn
		   (setq pm p1
			  p1 p2
			  p2 pm
			  p3 p4
		   )
		)
	)
	(if (equal p2 p4 0.01)
		(progn
		   (setq pm p1
			  p1 p2
			  p2 pm
		   )
		)
	)
	;angle : Returns the angle from two points (in radians)
	;(polar base-point angle distance) => Creates a new point starting from base-point, moving at the specified angle and distance
	(setq p4 (polar p3 (angle p1 p2) (distance p1 p2))) 
	
	(setq pp1 p1)
	(setq pp2 p2)
	(setq pp3 p3)
	(setq pp4 p4)

	(rect)
	
	;fix :Converts the result to an integer by truncating the decimal part
	;	 (does NOT round — it simply cuts off decimals). 
	;	x0 = integer_part_of(Xcp2/(0.1×ech​​))
	;Example:	(fix 3.9) → 3
	;			(fix -3.9) → -3
	(setq x0  (fix  (/ (nth 0 cp2) (* 0.1 ech)) )) 
	(setq y0  (fix  (/ (nth 1 cp2) (* 0.1 ech)) ))
	(setq x0 (* (+ x0 1) (* 0.1 ech)))
	(setq y0 (* (+ y0 1) (* 0.1 ech)))
	(setq p0 (list x0 y0 ))
	
	(setq ang1 (angle p1 p2))
	(if (<= pi ang1)
		(progn
			(setq pm p1
			   p1 p2
			   p2 pm
			)
			(setq pm p3
			   p3 p4
			   p4 pm
			)
		)
	)
	(setq ang1 (angle p1 p2))
	(setq ang2 (angle p1 p3))
	(if (<= pi ang2)
		(progn
			(setq pm p1
			   p1 p3
			   p3 pm
			)
			(setq pm p2
			 p2 p4
			 p4 pm
			)
		)
	)
	(setq ang2 (angle p1 p3))
	(if (< ang2 ang1)
		(progn
			(setq pm ang1
			   ang1 ang2
			   ang2 pm
			)
		)
	)
	(setq ang3 (angle p4 p2))
	(setq ang4 (angle p4 p3))
	(if (< ang4 ang3)
		(progn
			(setq pm ang3
			   ang3 ang4
			   ang4 pm
			)
		)
	)

	(setq yy y0)
	(while (< yy (nth 1 cp1)) 
		(setq xx x0)
		(while (< xx (nth 0 cp1)) 
		 (setq p (list xx yy)) 
		 (if (and (< ang1 (angle p1 p)) (< (angle p1 p) ang2) (< ang3 (angle p4 p)) (< (angle p4 p) ang4) )
		   (progn
			 (setq pt1 (list xx (- yy size)))
			 (setq pt2 (list xx (+ yy size)))
			 (Add-Line-OnLayer ms pt1 pt2 "Limite_plan")
			 (setq pt1 (list (- xx size) yy ))
			 (setq pt2 (list (+ xx size) yy ))
			 (Add-Line-OnLayer ms pt1 pt2 "Limite_plan")
		   );progn
		 );if
		 (setq xx (+ xx step))
		);while
	   (setq yy (+ yy step))
	);while
	
	(setq yy y0)
	(setq xx x0)
	(while (< xx (nth 0 cp1)) 
		(setq p (list xx yy)) 
		(setq pt1 (list xx (nth 1 cp2)))
		(setq pt2 (list xx (nth 1 cp1)))
		(setq p (inters pp1 pp2 pt1 pt2))
		(if p ;if p is not null
			(progn
			 (setq pt1 (list (nth 0 p) (+ (nth 1 p) (* ech 0.003))))
			 (setq ptt1 (list (nth 0 p) (+ (nth 1 p) (* ech 0.0045))))
			 (setq pt2 (list (nth 0 p) (- (nth 1 p) (* ech 0.003))))
			 (if (and (<= ang1 (angle p1 pt2)) (<= (angle p1 pt2) ang2) (<= ang3 (angle p4 pt2)) (<= (angle p4 pt2) ang4) )
			   (progn
				 (setq pt1 pt2)
				 (setq ptt1 (list (nth 0 p) (- (nth 1 p) (* ech 0.0045)))) 
			   )
			 )
			 (Add-Line-OnLayer ms p pt1 "Limite_plan")
			 (setq x (nth 0 p))
			 (setq dist (rtos x 2 0))
			 ;(setq newPt (list (car ptt1) (+ 13 (cadr ptt1))))
			 ;(setq newPt (list (car newPt) (+ (cadr newPt) (/ (* 13.0 ech) 2500.0)) ))
			 (if (equal pt1 pt2)
			   (Add-Text-OnLayer ms dist ptt1 h (* pi 1.5) "Limite_plan" "ml") ;; pi*1.5 = 270°
			   (Add-Text-OnLayer ms dist ptt1 h (* pi 1.5) "Limite_plan" "mr")
			 )
			);progn
		);if
		 
		(setq pt1 (list xx (nth 1 cp2)))
		(setq pt2 (list xx (nth 1 cp1)))
		;inters : Returns the intersection of of two lines: Line 1: from pp3 to pp4;Line 2: from pt1 to pt2
		(setq p (inters pp3 pp4 pt1 pt2))
		(if p
			(progn
				(setq pt1 (list (nth 0 p) (+ (nth 1 p) (* ech 0.003))))
				(setq pt2 (list (nth 0 p) (- (nth 1 p) (* ech 0.003))))
				(if (and (<= ang1 (angle p1 pt2)) (<= (angle p1 pt2) ang2) (<= ang3 (angle p4 pt2)) (<= (angle p4 pt2) ang4) )
				 (setq pt1 pt2)
				)
				(Add-Line-OnLayer ms p pt1 "Limite_plan")
			);progn
		);if 
		(setq xx (+ xx (* ech 0.10)))
	);while

	(setq xx y0)
	(setq yy y0)
	
	(while (< yy (nth 1 cp1)) 
		(setq p (list xx yy)) 
		(setq pt1 (list (nth 0 cp1) yy))
		(setq pt2 (list (nth 0 cp2) yy ))

		(setq p (inters pt1 pt2 pp1 pp3))
		(if p
			(progn
			 (setq pt1 (list (+ (nth 0 p) (* ech 0.003)) (nth 1 p) ))
			 (setq pt2 (list (- (nth 0 p) (* ech 0.003)) (nth 1 p) ))
			 (setq ptt1 (list (- (nth 0 p) (* ech 0.0045)) (nth 1 p)))
			 (if (and (< ang1 (angle p1 pt2)) (< (angle p1 pt2) ang2) (< ang3 (angle p4 pt2)) (< (angle p4 pt2) ang4) )
			   (setq pt1 pt2)
			   (setq ptt1 (list (+ (nth 0 p) (* ech 0.0045)) (nth 1 p)))
			 )
			 (Add-Line-OnLayer ms p pt1 "Limite_plan")
			 (setq y (nth 1 p))
			 (setq dist (rtos y 2 0))
			 (if (equal pt1 pt2)
			   (Add-Text-OnLayer ms dist ptt1 h 0.0 "Limite_plan" "mr") ; 0.0 = 0°
			   (Add-Text-OnLayer ms dist ptt1 h 0.0 "Limite_plan" "ml")
			 )
			);progn
		);if
		
		(setq pt1 (list (nth 0 cp1) yy))
		(setq pt2 (list (nth 0 cp2) yy ))
		(setq p (inters pt1 pt2 pp2 pp4))
		(if p
			(progn
				(setq pt1 (list (+ (nth 0 p) (* ech 0.003)) (nth 1 p) ))
				(setq pt2 (list (- (nth 0 p) (* ech 0.003)) (nth 1 p) ))
				(if (and (< ang1 (angle p1 pt2)) (< (angle p1 pt2) ang2) (< ang3 (angle p4 pt2)) (< (angle p4 pt2) ang4) )
				(setq pt1 pt2)
				)
				(Add-Line-OnLayer ms p pt1 "Limite_plan")
			);progn
		);if

		(setq yy (+ yy step))
	 );while

	(princ)
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Get cp1(MaxX;MaxY) cp1(MinX;MinY)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun rect ( / xs ys)
  (setq xs (mapcar 'car (list p1 p2 p3 p4)))
  (setq ys (mapcar 'cadr (list p1 p2 p3 p4)))

  (setq cp1 (list (apply 'max xs) (apply 'max ys)))
  (setq cp2 (list (apply 'min xs) (apply 'min ys)))
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Return Variants containing SafeArrays, not regular lists.
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun SafeArray->List (obj)
  (cond
    ((= (type obj) 'variant)
     (vlax-safearray->list (vlax-variant-value obj))
    )
    ((= (type obj) 'safearray)
     (vlax-safearray->list obj)
    )
  )
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Converts a 3D point list (x y z) into a COM Variant containing a SafeArray
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun Point->Variant (pt / sa)
  (setq sa (vlax-make-safearray vlax-vbDouble '(0 . 2)))
  (vlax-safearray-fill sa
    (list
      (car pt)
      (cadr pt)
      (if (caddr pt) (caddr pt) 0.0)
    )
  )
  (vlax-make-variant sa)
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Add Line on layer
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun Add-Line-OnLayer (ms p1 p2 layer / obj)
  (setq obj
    (vla-addLine ms
      (Point->Variant p1)
      (Point->Variant p2)
    )
  )
  (vla-put-Layer obj layer)
  (vla-put-Color obj 7)
  obj
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Add text on layer
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun Add-Text-OnLayer (ms txtStr pt height rot layer justify / txt)

	(setq txt (vla-AddMText ms (Point->Variant pt) 0 txtStr))

	;; height
	(vla-put-Height txt height)

	;; set final alignment
	(if (= justify "mr")
		(vla-put-AttachmentPoint txt acAttachmentPointMiddleRight)
		(vla-put-AttachmentPoint txt acAttachmentPointMiddleLeft)
	)
   
	;; reset insertion point (important!)
	(vla-put-InsertionPoint txt (Point->Variant pt))

	;; other properties
	(vla-put-Rotation txt rot)
	(vla-put-Layer txt layer)
	(vla-put-Color txt 7)
	(vla-put-StyleName txt "ITALIC")

	txt
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Add Layer, Auto-Creation + Add lintype
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
; (defun Ensure-Layer (doc name / layers)
  ; (setq layers (vla-get-Layers doc))
  ; (if (not (tblsearch "LAYER" name))
    ; (vla-add layers name)
  ; )
; )
(defun Ensure-Layer (doc name / layers)
  (setq layers (vla-get-Layers doc))
  (if (vl-catch-all-error-p
        (vl-catch-all-apply 'vla-item (list layers name)))
    (vla-add layers name)
  )
)

; (defun Ensure-Linetype (doc name file / ltypes lt)
  ; (setq ltypes (vla-get-Linetypes doc))

  ; (setq lt (vl-catch-all-apply 'vla-item (list ltypes name)))

  ; ;; If exists → delete it
  ; (if (not (vl-catch-all-error-p lt))
    ; (vla-delete lt)
  ; )

  ; ;; Reload from file
  ; (vla-load ltypes name file)
; )

(defun Ensure-Linetype (doc name file / ltypes) 
	(setq ltypes (vla-get-Linetypes doc)) 
	(if (vl-catch-all-error-p (vl-catch-all-apply 'vla-item (list ltypes name))) 
		(vla-load ltypes name file) 
	) 
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; ObjectDBX Document object, which allows manipulation of DWG files without opening a GUI drawing window.
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun _ObjectDBXDocument ( acapp / acVer )
	(vla-GetInterfaceObject acapp
	  (if (< (setq acVer (atoi (getvar "ACADVER"))) 16)
		"ObjectDBX.AxDbDocument" (strcat "ObjectDBX.AxDbDocument." (itoa acVer))
	  )
	)
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Detect and Skip 0-Byte DWG Files
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun is-empty-dwg (filename)
  (or
    (not (findfile filename)) ;File doesn't exist
    (= (vl-file-size filename) 0) ;File is 0 bytes
  )
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Detect Document is open by DWL
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun is-dwg-open-by-dwl (dwgpath pcName / dwl dwl2)
  (setq dwl  (strcat (vl-filename-base dwgpath) ".dwl"))
  (setq dwl2 (strcat (vl-filename-base dwgpath) ".dwl2"))

  (setq dwl  (strcat (vl-filename-directory dwgpath) "\\" dwl))
  (setq dwl2 (strcat (vl-filename-directory dwgpath) "\\" dwl2))
  
  (if (findfile dwl)
	(progn
      (setq file (open dwl "r"))
      (setq content "")
      (while (setq line (read-line file))
        (setq content (strcat content line "\n"))
      )
      (close file)
      (cond
		((null content)
			nil
		)
		((wcmatch (strcase content) (strcase (strcat "*" pcName "*")))
			T
		)
		(T
			nil
		)
	  )
    )
	
	(progn 
		(if (findfile dwl2)
			T
			nil
		)
	)
  )
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Start function Open csv Log file
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun _OpenFile ( fn / shell result )
	(setq result
	  (vl-catch-all-apply
		(function
		  (lambda nil
			(setq shell (vla-getInterfaceObject (vlax-get-acad-object) "Shell.Application"))
			(vlax-invoke shell 'open fn)
		  )
		)
	  )
	)
	(if shell (vlax-release-object shell))
	(not (vl-catch-all-error-p result))
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Start function to open folder dialogue to select the path
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun GetFolder (msg / Dir Item Path)
  (cond
    ((setq Dir (vlax-invoke (vlax-get-or-create-object "Shell.Application")
			    'browseforfolder
			    0
			    msg
			    1
			    ""
	       )
     )
     (cond
       ((not (vl-catch-all-error-p (vl-catch-all-apply 'vlax-invoke-method (list Dir 'Items))))
	(setq Item (vlax-invoke-method (vlax-invoke-method Dir 'Items) 'Item))
	(setq Path (vla-get-path Item))
	(if (not (member (substr Path (strlen Path) 1) (list "/" "\\")))
	  (setq Path (strcat Path "\\"))
	);end if
       )
     );end cond
    )
  );end cond
  Path
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Start function to Purge using Object DBX
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun Purge-DBX (dbxDoc / try-delete defcoll)

  ;Safe delete wrapper
  (defun try-delete (obj)
    (vl-catch-all-apply 'vla-delete (list obj))
  )

  ;Helper to purge a definition collection
  (defun defcoll (collection skiplist)
    (vlax-for item collection
      (if (not (member (strcase (vla-get-name item)) skiplist))
        (try-delete item)
      )
    )
  )

  ;Purge Blocks
  (defcoll (vla-get-blocks dbxDoc) '("MODEL_SPACE" "PAPER_SPACE"))

  ;Purge Layers (except 0 and Defpoints)
  (defcoll (vla-get-layers dbxDoc) '("0" "DEFPOINTS"))

  ;Purge Linetypes
  (defcoll (vla-get-linetypes dbxDoc) '("BYLAYER" "BYBLOCK" "CONTINUOUS"))

  ;You can add more here: dimension styles, text styles, etc.
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Load line Type "acadiso.lin"
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun LoadLinetypeToDBX (dbxDoc ltype)
	(if (not (tblsearch "LTYPE" ltype))
		(vla-load (vla-get-Linetypes dbxDoc) ltype "acadiso.lin")
	)
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Start Function Replace using Regex
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(or (vl-bb-ref '*REX*)
    (vl-bb-set '*REX* (vlax-create-object "VBScript.RegExp"))
)

  (defun _RegExReplace ( newstr pat string )
    (vlax-put    (vl-bb-ref '*REX*) 'Pattern pat)
    (vlax-put    (vl-bb-ref '*REX*) 'Global actrue)
    (vlax-put    (vl-bb-ref '*REX*) 'IgnoreCase acfalse)
    (vlax-invoke (vl-bb-ref '*REX*) 'Replace string newstr)
  )
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; FirstVertexOnContour
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun FirstVertexOnContour (pt ms / hit)
  (setq hit nil)

  (vlax-for ent ms
    (if (and
          (= (vla-get-objectname ent) "AcDbPolyline")
          (= (strcase (vla-get-layer ent)) "0")
		  (= (vla-get-Color ent) 1)
        )
      (if (< (distance pt (vlax-curve-getClosestPointTo ent pt)) 1e-6)
        (setq hit T)
      )
    )
  )
  hit
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; ModifyLengthFromP1
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun ModifyLengthFromP1 (ent scale / coords p1 p2 vec len newp2 newcoords)

  ;; read coordinates
  (setq coords
        (vlax-safearray->list
          (vlax-variant-value
            (vla-get-Coordinates ent))))

  ;; vertices
  (setq p1 (list (nth 0 coords) (nth 1 coords) 0.0))
  (setq p2 (list (nth 2 coords) (nth 3 coords) 0.0))

  ;; direction vector
  (setq vec (mapcar '- p2 p1))

  ;; current length
  (setq len (distance p1 p2))

  ;; normalize vector
  (setq vec (mapcar '(lambda (x) (/ x len)) vec))

  ;; compute new point
  (setq newp2
        (mapcar '+ p1
          (mapcar '(lambda (x) (* x scale)) vec)))

  ;; build new coordinates
  (setq newcoords
        (list
          (car p1) (cadr p1)
          (car newp2) (cadr newp2)))

  ;; write back coordinates
  (vla-put-Coordinates ent
    (vlax-make-variant
      (vlax-safearray-fill
        (vlax-make-safearray vlax-vbDouble '(0 . 3))
        newcoords)))

)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; ModifyLengthFromP2
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun ModifyLengthFromP2 (ent scale / coords p1 p2 vec len newp1 newcoords)

  ;; read coordinates
  (setq coords
        (vlax-safearray->list
          (vlax-variant-value
            (vla-get-Coordinates ent))))

  ;; vertices
  (setq p1 (list (nth 0 coords) (nth 1 coords) 0.0))
  (setq p2 (list (nth 2 coords) (nth 3 coords) 0.0))

  ;; vector from P2 to P1
  (setq vec (mapcar '- p1 p2))

  ;; current length
  (setq len (distance p1 p2))

  ;; normalize vector
  (setq vec (mapcar '(lambda (x) (/ x len)) vec))

  ;; compute new P1
  (setq newp1
        (mapcar '+ p2
          (mapcar '(lambda (x) (* x scale)) vec)))

  ;; build new coordinates
  (setq newcoords
        (list
          (car newp1) (cadr newp1)
          (car p2) (cadr p2)))

  ;; update polyline
  (vla-put-Coordinates ent
    (vlax-make-variant
      (vlax-safearray-fill
        (vlax-make-safearray vlax-vbDouble '(0 . 3))
        newcoords)))
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Split-MText-Lines
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
; (defun Split-MText-Lines (txt / lines)
  ; ;; Replace CRLF or LF with just LF
  ; (setq txt (vl-string-subst "\n" "\r\n" txt))
  
  ; ;; Split text into a list of lines
  ; (setq lines (vl-string->list txt "\n"))
  
  ; lines
; )
(defun Split-MText-Lines (txt / pos lines)
  (setq lines '())
  (while (setq pos (vl-string-search "\n" txt))
    (setq lines (cons (substr txt 1 pos) lines))
    (setq txt (substr txt (+ pos 2)))
  )
  (reverse (cons txt lines))
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Start Main Command Croisillon
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun c:Croisillon (/   ctf	  	  DwgPath    csvPathFile csvfile     pcName
							File      Files	     FilesList  i		 	     	
			   )
	
  ;Scripting.FileSystemObject : COM object provided by Microsoft that allows access to the file system operations
  (if (not FileSystemObject)
    (setq FileSystemObject
	   (vla-getInterfaceObject (vlax-get-acad-object) "Scripting.FileSystemObject")
    )
  );end if
  
  ;Initialize the counter i = 0
  (setq i 0)
  (setq pcName (getenv "COMPUTERNAME"))
  ;Start Condition
  (cond
    ((setq DwgPath (GetFolder "Selectionnez le dossier qui contient les fichiers autocad :"))
	 (setq outputPath (GetFolder "Selectionnez le dossier de sortie :"))
	 ;Set the path of log.csv file
	 (setq csvPathFile (strcat outputPath "\log.csv"))
	 ;Open the csv file for writing
	 (setq csvfile (open csvPathFile "w"))
	 ;the Header of csv :
	 (write-line "Filename,Probleme,Count" csvfile)
	 
	 ;Get All DWG files
     (setq Files (mapcar '(lambda (x) (strcat dwgpath x)) (vl-directory-files DwgPath "*.dwg" 1)))
	 
	 ;Start Process ...
     (Prompt "\n Starting process ...\n")
     (cond
		(Files
		(setq ctf (length files))
		
		(setq 	acadApp (vlax-get-acad-object) ; Get AutoCAD Application object
				dbxDoc (_ObjectDBXDocument acadApp) ; Returns a reference to a background DWG document that can be opened and manipulated without showing it in AutoCAD.
		)
		
		;Start Foreach Loop
		(foreach dwgFile Files
			(setq filename (vl-filename-base dwgFile)) ; Extract filename without path and extension
			(setq Ech 0)
			(if (is-empty-dwg dwgFile)
				(write-line (strcat filename ",Document invalid.") csvfile)
				(progn
					(if (not (is-dwg-open-by-dwl dwgFile pcName))
					(progn
						(if (/= (logand (vlax-get-property (vlax-invoke-method FileSystemObject 'getfile dwgFile) 'Attributes ) 1) 1)
						(progn
							(setq open-result 
							  (vl-catch-all-apply 'vlax-invoke-method (list dbxDoc 'Open dwgFile))
							)
							(if (vl-catch-all-error-p open-result)
							  (write-line (strcat filename ",incompatible version or Document is locked.") csvfile)
							  (progn
								;Start Modify the DWG
								(setq ms (vla-get-modelspace dbxDoc))
								
								;; Ensure the linetype is loaded - ACAD_ISO03W100
								; (if (not (tblsearch "LTYPE" "HIDDEN"))
									; (vla-load (vla-get-Linetypes dbxDoc) "HIDDEN" "acadiso.lin")
								; )
								;(Ensure-Linetype dbxDoc "HIDDEN" "acadiso.lin")
								(Ensure-Linetype dbxDoc "DASHED" "acadiso.lin")
								
								;;Creat layer if not exist
								(Ensure-Layer dbxDoc "Limite_plan")
								(Ensure-Layer dbxDoc "PDF_TEXT_PLAN_SCALE")
								
								;1 Vlax-For Detect Scale,N° demande and Nom
							    (vlax-for ent ms
								  (if (and (= (vla-get-objectname ent) "AcDbMText")
										   (= (vla-get-Layer ent) "PDF_TEXT_PLAN_SCALE")) ; Contents : {\C256;\c16777215;1:1 500}
									(if (vl-string-search "1:" (vla-get-TextString ent))
									  (progn
										(setq Ech (vla-get-TextString ent))
										(setq Ech (substr Ech (+ (vl-string-search ";" Ech 2) 1)))
										(setq Ech (substr Ech (+ (vl-string-search ";" Ech 2) 1)))
										
										(setq Ech (vl-string-subst "" "}" Ech))   ; Remove "}"
										(setq Ech (vl-string-subst "" ";" Ech))   ; Remove ";"
										; Remove the "1:" and the space
										(setq Ech (vl-string-subst "" "1:" Ech))  ; Remove "1:"
										(setq Ech (vl-string-subst "" " " Ech))   ; Remove the space
										
										; Convert the cleaned string to an integer
										(setq Ech (atoi Ech))  ; Convert to integer
									  )
									)
								  )
								  ; (if (and (= (vla-get-objectname ent) "AcDbMText")
										   ; (= (vla-get-Layer ent) "Text_info"))
									; (if (vl-string-search "dossier de demande :" (vla-get-TextString ent))
									  ; (progn
										; (setq num_demande (vla-get-TextString ent))
										; ; Remove the "N° dossier de demande :"
										; (setq pos (vl-string-search ": " num_demande))
										; (setq num_demande (substr num_demande (+ pos 2)))
										; ; in one line
										; ;(setq num_demande (substr Ech (+ (vl-string-search ":" Ech) 2)))
										; (setq num_demande (vl-string-subst "" " " num_demande)); Remove the space
									  ; )
									; )
								  ; )
								  (if (and (= (vla-get-objectname ent) "AcDbMText")
										   (= (vla-get-Layer ent) "0")
										   (vl-string-search "Parcelle de" (vla-get-TextString ent))
										)
									(progn
										(setq txt (vla-get-TextString ent))
										(setq txt (vl-string-subst "\n" "\\P" txt))
										(setq txt (vl-string-subst "\n" "\\P" txt))
										(setq lines (Split-MText-Lines txt))
										;; Now you can assign each line to a variable
										(setq line1 (nth 0 lines)) ;line1
										(setq owners (nth 1 lines)) ;line2
										(setq num_demande (nth 2 lines)) ;line3
									)
								  )
								)
								(princ)
								
								;2 Vlax-For 
							    (vlax-for ent ms
								  (if (and (= (vla-get-objectname ent) "AcDbMText")
										   (= (vla-get-Layer ent) "Text_info")
										   (vlax-property-available-p ent 'TextString)
										)
									(progn
										(if (vl-string-search "dossier de demande :" (vla-get-TextString ent))
										  (progn
											(vla-put-TextString ent (strcat "N%%d dossier de demande : " num_demande))
										  )
										)
										(if (vl-string-search "Nom du demandeur : " (vla-get-TextString ent))
										  (progn
											(vla-put-TextString ent (strcat "Nom du demandeur : " owners))
										  )
										)
									)
								  )
								)
								(princ)
								
								(write-line (strcat filename "," (if Ech (Itoa Ech) "") "," (if num_demande num_demande "")) csvfile)
								
								;3 Vlax-For Detect Scale
							    (vlax-for ent ms
								  (if (and (= (vla-get-objectname ent) "AcDbMText")
										   (= (vla-get-Layer ent) "PDF_TEXT_PLAN_SCALE"))
									(if (vl-string-search "1:" (vla-get-TextString ent))
									  (progn
										(setq ins (vlax-get ent 'InsertionPoint))
										(setq str (vla-get-TextString ent))
										(setq hgt (vla-get-Height ent))
										(setq newEch (strcat "1:" (Itoa Ech)))
										(Add-Text-OnLayer ms newEch ins hgt 0.0 "PDF_TEXT_PLAN_SCALE" "ml")
										(vla-delete ent)
									  )
									)
								  )
								)
								(princ)
								 
								(setq scaleLT (/ Ech 1000.0))
								(setq scale (* 20.0 (/ Ech 1000.0)))
								
								;4 Vlax-For
								(vlax-for ent ms
									(if (and 	(= (vla-get-objectname ent) "AcDbMText")
												(= (vla-get-Layer ent) "0")
											)
										(progn
											(vla-put-Color ent 7)
											(if (vl-string-search "P" (vla-get-TextString ent))
												(progn
													(setq modifiedText (_RegExReplace "Parcelle de" "(^|\r?\n)P?\\d+" (vla-get-TextString ent)) ) ;(^|\r?\n)P?\\d+ ;;\\bP[0-9]+\r?\n?
													(vla-put-TextString ent modifiedText)
												)
											)
										);end progn
									);end if
									
									(if (and (= (vla-get-objectname ent) "AcDbMText")
										   (= (vla-get-Layer ent) "Sommair_map")
										   (vl-string-search "Village de " (vla-get-TextString ent))
										)
										(vla-put-TextString ent "Parcelle concern\U+00E9e") ;;"\U+00E9" → é "\U+00E8" → è "\U+00E0" → à "\U+00E7" → ç
									)
								  
									(if (and 	(= (vla-get-objectname ent) "AcDbPolyline")
												(= (vla-get-Layer ent) "MAP_DETAIL_LINE")
											)
										(progn
											(vla-put-Color ent 1)
										);end progn
									);end if
									
									(if (and
											(= (vla-get-objectname ent) "AcDbPolyline")
											(= (vla-get-layer ent) "RIVRAIN_SEGMENTS")
										)
									(progn
										(vla-put-Linetype ent "DASHED") ;ByLayer or HIDDEN
										(vla-put-LinetypeScale ent scaleLT)
										; (if (vlax-property-available-p ent 'LinetypeGeneration)
											; (vla-put-LinetypeGeneration ent :vlax-true)
										; )
										(setq coords
											(vlax-safearray->list
											  (vlax-variant-value (vla-get-Coordinates ent))
											)
										)

										;; check 2 vertices
										(if (= (length coords) 4)
											(progn
											  (setq p1 (list (nth 0 coords) (nth 1 coords) 0.0))
											  (setq p2 (list (nth 2 coords) (nth 3 coords) 0.0))

											  (if (FirstVertexOnContour p1 ms)
												;; normal
												(ModifyLengthFromP1 ent scale)
												;; reverse
												(ModifyLengthFromP2 ent scale)
											  )
											)
										)
									)
								  )
								 
								);End Vlax for
								(princ)
								
								;5 Vlax-For Detect Polyline & Delete Old Cross
								(vlax-for ent ms
									(cond
										((and (= (vla-get-ObjectName ent) "AcDbPolyline")
											  (= (vla-get-Layer ent) "Limite_plan"))
										 (setq pl ent)
										)

										((and (member (vla-get-ObjectName ent) '("AcDbLine" "AcDbMText"))
											  (= (vla-get-Layer ent) "Limite_plan"))
										 (vla-delete ent)
										)
									)
								)
								(princ)
								
								(if (and (numberp Ech) (> Ech 0))
									(Create-Croix ms Ech pl)
									(write-line (strcat filename ",Echelle introuvable.") csvfile)
								)
								
								;; Run manual purge
								(Purge-DBX dbxDoc)

								(setq i (1+ i));iterate the Counter
								(setq newfilename (strcat "CF-" (if num_demande num_demande "") "-PPDef"))
								;Saveas
								(vla-saveas dbxDoc (strcat outputPath newfilename ".dwg"))
							  )
							)
						)
						(write-line (strcat filename ",is read only.") csvfile)
						)
					)
					(write-line (strcat filename ",is open now.") csvfile)
					)
				)
			)
		);end foreach
		(if dbxDoc (vlax-release-object dbxDoc)) ;release the COM object created by dbxDoc
		(Gc) ;Garbage Collection : explicitly triggers garbage collection
		(Gc) ; identify and reclaim memory occupied by objects that are no longer reachable or in use by the program
		(Gc)
       )
       (T (Prompt "\nNothing files found."))
     );end cond
	 ;Colse the csv file
	 (close csvfile)
	 ; Open the CSV file (Log.csv)
	 (if (null (_OpenFile csvPathFile))
		(princ "\n--> Error Opening Report.")
		(princ "\n--> CSV Report Opened.")
	  )
    )
    (T (Prompt "\nNothing selected. "))
  );end condition
  
  
  (princ (Strcat "\n DONE. Processed " (Itoa i) " drawings. !!"))
  (princ)
) ;end Croisillon
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(vl-load-com);Load the COM object
(princ)

(princ "\n Lisp Loaded Correctly.")
(princ "\n Let's Start to use Croisillon Command :)")
