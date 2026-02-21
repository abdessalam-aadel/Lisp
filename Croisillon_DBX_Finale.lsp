;;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Croisillon_DBX_Tonkpi : Small lisp to Creat quadrillage ... in Multiple DWG file
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
	
	;; Start Bounding box of polyligne
	(vla-getBoundingBox pl 'minPt 'maxPt)

	(setq p2 (SafeArray->List minPt))
	(setq pp4 (SafeArray->List maxPt))
	(setq p1 (list (car pp4) (cadr p2)))
	(setq p3 p2)
	(setq p4 (list (car p2) (cadr pp4)))
	
	(setq step (* ech 0.10))
	(setq size (* ech 0.0025))
	(setq h (* 0.002 ech))
	
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
	
	(setq p4 (polar p3 (angle p1 p2) (distance p1 p2))) ; (polar base-point angle distance)
	
	(setq pp1 p1)
	(setq pp2 p2)
	(setq pp3 p3)
	(setq pp4 p4)

	(rect)
   
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
			 (setq newPt (list (car ptt1) (+ 13 (cadr ptt1))))
			 (Add-Text-OnLayer ms dist newPt h (* pi 1.5) "Limite_plan") ;; pi*1.5 = 270°
		   );progn
		 );if
		 
		(setq pt1 (list xx (nth 1 cp2)))
		(setq pt2 (list xx (nth 1 cp1)))
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
			 (Add-Text-OnLayer ms dist ptt1 h 0.0 "Limite_plan") ; 0.0 = 0°
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
;;; SafeArray
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
;;; Point to SafeArray
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
  obj
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Add text on layer
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun Add-Text-OnLayer (ms txtStr pt height rot layer / txt)
  (setq txt (vla-AddText ms txtStr (Point->Variant pt) height))
  (vla-put-Alignment txt acAlignmentMiddleLeft)
  (vla-put-TextAlignmentPoint txt (Point->Variant pt))
  (vla-put-Rotation txt rot)
  (vla-put-StyleName txt "ITALIC")
  (vla-put-Layer txt layer)
  txt
)
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Add Layer Auto-Creation
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun Ensure-Layer (doc name / layers)
  (setq layers (vla-get-Layers doc))
  (if (not (tblsearch "LAYER" name))
    (vla-add layers name)
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
								
								;;Creat layer if not exist
								(Ensure-Layer dbxDoc "Limite_plan")
								
								;1 Vlax-For Detect Scale
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
										
										(write-line (strcat filename "," Ech) csvfile)
										; Convert the cleaned string to an integer
										(setq Ech (atoi Ech))  ; Convert to integer
									  )
									)
								  )
								)
								(princ)
								
								;2 Vlax-For Detect Polyline & Delete Old Cross
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

								(setq i (1+ i));iterate the Counter
								
								;Saveas
								(vla-saveas dbxDoc (strcat outputPath filename ".dwg"))
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
