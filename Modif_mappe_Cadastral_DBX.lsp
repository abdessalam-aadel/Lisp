;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Modif_mappe_Cadastral Lisp : is a samll Routine to Delete Layer (RECT) and layer (crois)
;;; and Delete the Mtext contain "Echelle"
;;; or contain the Mappe ex : 26-9-1-m in Multiple DWG Files and Purge them, the result in c:\output folder
;;;
;;; Copyright © 2025
;;; https://github.com/abdessalam-aadel/Lisp
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
;;; Start Main Command ModifMappe_DBX 
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun c:ModifMappe_DBX (/   ctf	 DwgPath     outputPath   i
							File     Files	     FilesList	  csvPathFile	 	 csvfile     	
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
	 (write-line "Filename,Probleme" csvfile)
	 
	 ;Get All DWG files
     (setq Files (mapcar '(lambda (x) (strcat dwgpath x)) (vl-directory-files DwgPath "*.dwg" 1)))
	 
	 ;Start Process ...
     (Prompt "\n Starting process ...\n")
     (cond
		(Files
		(setq ctf (Length files))
		
		(vlax-for & (vla-get-documents (vlax-get-acad-object))
		  (setq FilesList (cons (strcase (vla-get-fullname &)) FilesList))
		)
		
		(setq 	acadApp (vlax-get-acad-object) ; Get AutoCAD Application object
				dbxDoc (_ObjectDBXDocument acadApp) ; Returns a reference to a background DWG document that can be opened and manipulated without showing it in AutoCAD.
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
			  ( (setq filename (vl-filename-base &)) ; Extract filename without path and extension
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
					)
					
					(vlax-for ent ms
						(cond
						  ;; 1. Delete entities on layer "Echelle"
						  ((and (= (vla-get-objectname ent) "AcDbMText")    
								(vl-string-search "Echelle" (vla-get-TextString ent))
							)
								(vla-delete ent)
						  )

						  ;; 2. Delete entities on layer "crois" and "RECT"
						  ((or (= (vla-get-Layer ent) "crois")				  
								(= (vla-get-Layer ent) "RECT")
							)
								(vla-delete ent)
						  )
						  
						  ;; 3. Delete Mtext contain filename
						  ((and (= (vla-get-objectname ent) "AcDbMText")    
								(vl-string-search filename (vla-get-TextString ent))
							)
								(vla-delete ent)
						  )
						)
					)
					(princ)
					(setq layer (vla-item layers "contour_req"))  ; Get the layer by name
					(vla-put-color layer 1)
					
				   ;End Modify the Mappe
				   
				   ;(Prompt (Strcat "\n Purge " & ". Please wait..."))
				   ;; Run manual purge
				   (Purge-DBX dbxDoc)
				   
				   ;Save & Close & Release Object
				   ;(Prompt (Strcat "\n Save and Close " & "\n"))
				   (vla-saveas dbxDoc (strcat outputPath filename ".dwg"))
				   (setq i (1+ i));iterate the Counter
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
) ;end ModifMappe_DBX
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(princ)

(princ "\n Lisp Loaded Correctly.")
(princ "\n Let's Start to use ModifMappe_DBX Command :)")