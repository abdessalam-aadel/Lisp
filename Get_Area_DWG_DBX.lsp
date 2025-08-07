;;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; GETAREA Lisp : is a samll Routine to get area of polyligne in layer "contour_"
;;; in Multiple DWG Files
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
;;; Start Main Command getarea
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun c:getarea (/   ctf	  	  DwgPath    csvPathFile csvfile     acadApp dbxDoc
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
  
  ;Start Condition
  (cond
    ((setq DwgPath (GetFolder "Selectionnez le dossier qui contient les fichiers autocad :"))
	 ;Set the path of GetArea.csv file
	 (setq csvPathFile (strcat DwgPath "\GetArea.csv"))
	 ;Open the csv file for writing
	 (setq csvfile (open csvPathFile "w"))
	 ;the Header of csv :
	 (write-line "Filename,Area" csvfile)
	 
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
					;Start Modify the DWG
					(setq ms (vla-get-modelspace dbxDoc)
						 layers (vla-get-Layers dbxDoc) ; Get the layers collection
					)
					
					;; Start processing all polylines
					(vlax-for ent ms
					 (if (and (= (vla-get-objectname ent) "AcDbPolyline")    
						 (= (vla-get-Layer ent) "contour_"))
						  (progn
							(setq area (vla-get-Area ent))
							(write-line (strcat filename "," (rtos area 2 5)) csvfile)
						   )
					 );end if
					);End Vlax for
					(princ)
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
		(Gc) ;identify and reclaim memory occupied by objects that are no longer reachable or in use by the program
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
) ;end getarea
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(princ)

(princ "\n Lisp Loaded Correctly.")
(princ "\n Let's Start to use getarea Command :)")