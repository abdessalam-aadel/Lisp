;;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; CorrigerPlan - Clean & Process DWG files via ObjectDBX
;;; - Set all layers to white (ACI color 7)
;;; - Delete all polylignes in layer "contour_" with Area = 0.0
;;; - Set all entities' color to ByLayer
;;; - Move closed polylines from "public.parcelle" to "contour_"
;;; - Bring "contour_" polylines to front
;;; - Purge unused elements (blocks, layers, linetypes)
;;; - Process multiple DWG files in batch, write log to CSV
;;;
;;; Author: Abdessalam Aadel © 2025
;;; GitHub: https://github.com/abdessalam-aadel/Lisp
;;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
;;; Start function to set the color object to By layer
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun dsx-put-color (obj num / av)
  (setq av (substr (getvar "acadver") 1 2))
  (if (>= av "16")
    ;; if AutoCAD 2004...
    (dsx-put-color2004 obj num) 
    ;; if any other version...
    (dsx-put-property obj "Color" num)
  )
)

(defun dsx-put-color2004 (obj num / oColor numlst) 
  (if
    (not 
      (vl-catch-all-error-p
        (setq oColor
          (vl-catch-all-apply 'vla-get-TrueColor (list obj))
        ) 
      )
    )
    (progn ;; if getting the TrueColor object of 'obj' did not return a vla-error
      (cond
        ( (= 'INT (type num)) ;; if an ACI index integer is passed 
          ;; if obj is a table record (i.e. Layer)...
          (if (vl-string-search "Table" (vla-get-ObjectName obj))
            (progn
              (vla-put-ColorMethod oColor acColorMethodByACI) 
              (vla-put-ColorIndex oColor num)
            )
            (if (= num acBylayer) ;; if obj is an entity
              (progn ;; if num is to be byLayer 
                (vla-put-ColorMethod oColor acColorMethodByLayer)
                (vla-put-ColorIndex oColor acByLayer)
              )
              (progn ;; if num is to be an override 
                (vla-put-ColorMethod oColor acColorMethodByACI)
                (vla-put-ColorIndex oColor num)
              )
            ) 
          )
        )
        ( (and (listp num) (= (length num) 3)) ;; an RGB list is passed
          (vla-put-ColorMethod oColor acColorMethodByRGB) ;; set the method 
          ;; set the RGB values
          (vlax-invoke-method oColor 'SetRGB (nth 0 num) (nth 1 num) (nth 2 num)) 
        )
      )
      ;; stuff color object back into parent object 
      (vla-put-TrueColor obj oColor)
      ;; clean up the memory stack of unused objects
      (vlax-release-object oColor)
    ) 
    (vl-catch-all-error-message oColor)
  ) 
)

(defun dsx-put-property (obj prop val / try)
  (cond 
    ( (and
        (vlax-property-available-p obj prop) 
        (not
          (vl-catch-all-error-p
            (setq try 
              (vl-catch-all-apply 'vlax-put-property (list obj prop val))
            )
          )
        ) 
      )
      val
    )
  ) 
)

;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Start function to Purge using Object DBX
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun Purge-DBX (dbxDoc / try-delete defcoll)

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
;;; Start Main Command CorrigerPlan2
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun c:CorrigerPlan (/   ctf	  	  DwgPath    csvPathFile csvfile
							File      Files	     FilesList	 outputPath  i		 	     	
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
		
		(vlax-for & (vla-get-documents (vlax-get-acad-object))
		  (setq FilesList (cons (strcase (vla-get-fullname &)) FilesList))
		)
		
		(setq 	acadApp (vlax-get-acad-object) ; Get AutoCAD Application object
				dbxDoc (_ObjectDBXDocument acadApp) ; Returns a reference to a background DWG document that can be opened and manipulated without showing it in AutoCAD.
		)
		
		; (setq Express
			; (and (vl-position "acetutil.arx" (arx))
			  ; (not
				; (vl-catch-all-error-p
				  ; (vl-catch-all-apply
					; (function (lambda nil (acet-sys-shift-down)))
				  ; )
				; )
			  ; )
			; )
		; )
		;Start progress bar
		; (if Express
            ; (setq ProgBar (acet-ui-progress "Start Modify the DWG..." ctf))
		; )
	
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
			  ( (setq filename (vl-filename-base &)) ; Extract filename without path and extension
				;(if Express (acet-ui-progress -1))
				
				(setq open-result 
				  (vl-catch-all-apply 'vlax-invoke-method (list dbxDoc 'Open &))
			    )
			   (if (vl-catch-all-error-p open-result)
				  (progn
					(write-line (strcat filename ",Drawing file was created by an incompatible version.") csvfile)
					(vlax-release-object dbxDoc)
				  )
				  (progn
					;(Prompt (Strcat "\n Open " & ". Please wait (" (Itoa (1+ i)) "/" (Itoa ctf) ")..."))
					;Start Modify the DWG
					(setq ms (vla-get-modelspace dbxDoc)
						 layers (vla-get-Layers dbxDoc) ; Get the layers collection
						 countContour 0 ; nbr of Polyline in layer "contour_"
					)
					
					;; Loop through all layers
					(vlax-for layer layers
						(if (not (equal (vla-get-Name layer) "contour_"))
							(vla-put-Color layer 7) ;Set the color to white (ACI color index 7)
						)
					)
					(princ)
					
					;; Loop through all objects
					(vlax-for ent ms
						(if (and (= (vla-get-ObjectName ent) "AcDbPolyline")
								 (= (vla-get-Layer ent) "contour_")
								 (= (vla-get-Area ent) 0.0)
							) 
								(vla-Delete ent)
						)
					)
					(princ)
					
					;; Loop through all objects
					(vlax-for ent ms
						(dsx-put-color ent acByLayer)
						(if (and (= (vla-get-ObjectName ent) "AcDbPolyline")
								 (= (vla-get-Layer ent) "public.parcelle")
								 (= (vla-get-Closed ent) :vlax-true)
							)
							(progn
								(vla-put-Layer ent "contour_")
							)
						)
						
						(if (and (= (vla-get-ObjectName ent) "AcDbPolyline")
								 (= (vla-get-Layer ent) "contour_")
							)
							(progn 
								(setq new (vla-Copy ent))
								(vla-Delete ent) ; delete the old one
								new ; now the new copy is on top
							)
						)
					)
					(princ)
				    ;End Modify the Mappe
				   
				    ;(Prompt (Strcat "\n Purge " & ". Please wait..."))
				    ;; Run manual purge
				    (Purge-DBX dbxDoc)
				   
				    ;; Loop through all objects
					(vlax-for ent ms
						(if (and (= (vla-get-ObjectName ent) "AcDbPolyline")
								 (= (vla-get-Layer ent) "contour_")
							) 
								(setq countContour (1+ countContour));iterate the Counter of polyline
						)
					)
					(princ)
				   
					;Save & Close & Release Object
					;(Prompt (Strcat "\n Save and Close " & "\n"))
					(vla-saveas dbxDoc (strcat outputPath filename ".dwg"))
					(setq i (1+ i));iterate the Counter
					;(setq countContour (- countContour 1))
					(if (> countContour 1)
						(write-line (strcat filename ",Nombre de polylignes dans le layer contour_ est :," (itoa countContour)) csvfile)
					)
				  )
				)
			   )
			  (T (write-line (strcat filename ",Drawing file was created by an incompatible version.") csvfile))
			);end cond
		    )
				(T (write-line (strcat filename ",is read-only.") csvfile))
			);end cond
			)
			(T (write-line (strcat filename ",is open now.") csvfile))
		  );end cond
		);end foreach
		;End progress bar
		;(if Express (setq ProgBar (acet-ui-progress)))
		(vlax-release-object dbxDoc)
		(Gc) ;Garbage Collection : explicitly triggers garbage collection
		(Gc) ; identify and reclaim memory occupied by objects that are no longer reachable or in use by the program
		(Gc)
       )
       (T (write-line "\nNothing files found to purge." csvfile))
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
) ;end CorrigerPlan
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(princ)

(princ "\n Lisp Loaded Correctly.")
(princ "\n Let's Start to use CorrigerPlan Command :)")