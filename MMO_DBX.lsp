;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; MMO Lisp : is a samll Routine to Delete Layer (titres) and Raster Image
;;; and Delete the Mtext contain "ZONE EXCLUS" and Mtext with color cyan and yellow, and Hachure the Mappe, an Purge
;;; in Multiple DWG Files the result in c:\output folder
;;;
;;; Copyright © 2025
;;; https://github.com/abdessalam-aadel/Lisp
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++

;; ObjectDBX Document object, which allows manipulation of DWG files without opening a GUI drawing window.
(defun _ObjectDBXDocument ( acapp / acVer )
	(vla-GetInterfaceObject acapp
	  (if (< (setq acVer (atoi (getvar "ACADVER"))) 16)
		"ObjectDBX.AxDbDocument" (strcat "ObjectDBX.AxDbDocument." (itoa acVer))
	  )
	)
)

(defun make-square-polyline (center ms / halfx halfy x y coords pline pts)
  ;; Calculate half-size
  (setq halfx (/ 0.8636 2.0))
  (setq halfy (/ 0.5758 2.0))
  (setq x (car center))
  (setq y (cadr center))

  ;; Define square corners (counter-clockwise)
  (setq pts (list
              (- x halfx) (- y halfy)
              (+ x halfx) (- y halfy)
              (+ x halfx) (+ y halfy)
              (- x halfx) (+ y halfy)
              (- x halfx) (- y halfy) ; back to start to close
            )
	)

  ;; Create the safearray
  (setq coords (vlax-make-safearray vlax-vbDouble (cons 0 (- (length pts) 1))))
  (vlax-safearray-fill coords pts)

  (setq pline (vla-AddLightWeightPolyline ms coords))

  ;; Close the polyline
  (vla-put-Closed pline :vlax-true)
  
  pline
)

;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Start function to Purge using Object DBX
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun Purge-DBX (dbxDoc / try-delete defcoll item)

  ;; Safe delete wrapper
  (defun try-delete (obj)
    (vl-catch-all-apply 'vla-delete (list obj))
  )

  ;; Helper to purge a definition collection
  (defun defcoll (collection skiplist)
    (vlax-for item collection
      (if (not (member (strcase (vla-get-name item)) skiplist))
        (try-delete item)
      )
    )
  )

  ;; Purge Blocks
  (defcoll (vla-get-blocks dbxDoc) '("MODEL_SPACE" "PAPER_SPACE"))

  ;; Purge Layers (except 0 and Defpoints)
  (defcoll (vla-get-layers dbxDoc) '("0" "DEFPOINTS"))

  ;; Purge Linetypes
  (defcoll (vla-get-linetypes dbxDoc) '("BYLAYER" "BYBLOCK" "CONTINUOUS"))

  ;; You can add more here: dimension styles, text styles, etc.
)

;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Start function to open folder dialogue to select the path
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun GetFolder (/ Dir Item Path)
  (cond
    ((setq Dir (vlax-invoke (vlax-get-or-create-object "Shell.Application")
			    'browseforfolder
			    0
			    "Select the Path with DWG files:"
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
;;; Start Main Command MMO_DBX
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun c:MMO_DBX (/   ctf	  	   DwgPath
			    File      Files	     FilesList	i	 SubDir	 	     	
			   )
	
  ;Scripting.FileSystemObject : COM object provided by Microsoft that allows access to the file system operations
  (if (not FileSystemObject)
    (setq FileSystemObject
	   (vla-getInterfaceObject (vlax-get-acad-object) "Scripting.FileSystemObject")
    )
  );end if
  
  ;Initialize the counter i = 0
  (setq i 0)
  
  ;Start Condition
  (cond
    ((setq DwgPath (GetFolder))
     
     (setq Files (mapcar '(lambda (x) (strcat dwgpath x)) (vl-directory-files DwgPath "*.dwg" 1)))
	 
     (Prompt "\n Starting process ...\n")
     (cond
		(Files
		(setq ctf (Length files))
		
		(vlax-for & (vla-get-documents (vlax-get-acad-object))
		  (setq FilesList (cons (strcase (vla-get-fullname &)) FilesList))
		)
	
		;Start Foreach Loop
		(foreach & Files
		  (cond
			((not (member & FilesList))
			 (cond
			   ((/= (logand (vlax-get-property (vlax-invoke-method FileSystemObject 'getfile &)
							   'Attributes
					)
					1
				)
				1
			)
			(cond
			  ((setq acadApp (vlax-get-acad-object) ; Get AutoCAD Application object
					dbxDoc (_ObjectDBXDocument acadApp)
				)
			  (Vlax-invoke-method dbxDoc 'Open &)
			   (setq filename (vl-filename-base &)) ; Extract filename without path and extension
			   (Prompt (Strcat "\n Open " & ". Please wait (" (Itoa (1+ i)) "/" (Itoa ctf) ")..."))
			   
				;Start Modify the DWG
				(Prompt "\n Start Modify the DWG ...")
				(setq ms (vla-get-modelspace dbxDoc)
					 textPos '(0.0 0.0 0.0)
				)
				
				(vlax-for ent ms
					(cond
					  ;; 1. Delete entities on layer "titres"
					  ((or  (= (vla-get-Layer ent) "titres")
							(= (vla-get-ObjectName ent) "AcDbRasterImage")
							(= (vla-get-ObjectName ent) "AcDbZombieEntity")
							(and (= (vla-get-ObjectName ent) "AcDbMText")
								(vl-string-search "ZONE" (vla-get-TextString ent))
								(vl-string-search "EXCLUS" (vla-get-TextString ent))
							)
							(and (= (vla-get-ObjectName ent) "AcDbMText")
								(or (= (vla-get-color ent) 4)
									(= (vla-get-color ent) 2)
								)
							)
						)
							(vla-delete ent)
					  )
					  
					  ;; 2. Explode block references named "1"
					  ((and (= (vla-get-ObjectName ent) "AcDbBlockReference")
							(or (= (vla-get-Name ent) "1")
								(= (vla-get-Name ent) "2")
								(= (vla-get-Name ent) "3")
								(= (vla-get-Name ent) "4")
								(= (vla-get-Name ent) "5")
							)
					   )
					   (vla-Explode ent)
					  )
					)
				)
				(princ)
				
				(vlax-for ent ms
					(cond
					  ;; Find MText on "BORNES" layer containing the filename
					  ((and (= (vla-get-ObjectName ent) "AcDbMText")
							(= (strcase (vla-get-Layer ent)) "BORNES")
							(vl-string-search filename (vla-get-TextString ent))
					   )
					   (setq textPos (vlax-get ent 'InsertionPoint))
					  )
					)
				)
				(princ)
				
				;; Create a temporary polyline using textPos (the center of polyline)
				(setq poly (make-square-polyline (list (car textPos) (cadr textPos)) ms))
				;; ---------------------------------------------
				;; Start Creating a hatch ----------------------
				;; ---------------------------------------------
				(setq hatch (vla-AddHatch
							 ms
							 acHatchPatternTypePreDefined
							 "ANSI31"
							 :vlax-true))
				;; Set hatch style (angle, scale)
				(vla-put-Color hatch 5)
				(vla-put-PatternAngle hatch 1.55) ; degrees
				(vla-put-PatternScale hatch 0.05)  ; smaller = tighter pattern

				;; Create a safearray for the loop
				(setq loopArray (vlax-make-safearray vlax-vbObject '(0 . 0)))
				(vlax-safearray-put-element loopArray 0 poly)

				;; Append the outer loop
				(vla-AppendOuterLoop hatch loopArray)

				;; Evaluate the hatch
				(vla-Evaluate hatch)
				(vla-delete poly) ; delete temporary polyline
				(princ "\n Hatch applied to enclosing polyline.")
				;; ---------------------------------------------
				;; End Creating a hatch ----------------------
				;; ---------------------------------------------
				
			   ;End Modify the Mappe
			   
			   ;Start Purge-All
			   (Prompt (Strcat "\n Purge " & ". Please wait..."))
			   ;; Run manual purge
				(Purge-DBX dbxDoc)
			   
			   ;Save & Close & Release Object
			   (Prompt (Strcat "\n Save and Close " & "\n"))
			   ;; if folder output not exist creat them
			   (if (not (vl-file-directory-p "C:\\output\\"))
				  (vl-mkdir "C:\\output\\")
				)
			   (vla-saveas dbxDoc (strcat "C:\\output\\" filename ".dwg"));(vla-saveas dbxDoc &) work fine with no .bak file
			   (vlax-release-object dbxDoc)
			   (setq i (1+ i));iterate the Counter
			  )
			  (T
			   (prompt (strcat "\nCannot open "
					   &
					   "\nDrawing file was created by an incompatible version. "
				   )
			   )
			  )
			);end cond
			   )
			   (T (prompt (strcat & " is read-only. Purge canceled. ")))
			 );end cond
			)
			(T (prompt (strcat & " is open now. Purge canceled. ")))
		  );end cond

		);end foreach
       )
       (T (Prompt "\nNothing files found to purge. "))
     );end cond
    )
    (T (Prompt "\nNothing selected. "))
  );end condition
  
  
  (princ (Strcat "\n DONE. Processed " (Itoa i) " drawings. !!"))
  (princ)
) ;end MMO_DBX
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(princ)

(princ "\n Lisp Loaded Correctly.")
(princ "\n Let's Start to use MMO_DBX Command :)")