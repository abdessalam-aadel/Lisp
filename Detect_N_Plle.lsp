;;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; DetectNPlle : Small lisp to detect Plle N° ... and ST .. in Multiple DWG file
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
;;; Start Main Command DetectNPlle
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun c:DetectNPlle (/   ctf	  	  DwgPath    csvPathFile csvfile     pcName
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
	 ;(setq outputPath (GetFolder "Selectionnez le dossier de sortie :"))
	 ;Set the path of log.csv file
	 (setq csvPathFile (strcat DwgPath "\log.csv"))
	 ;Open the csv file for writing
	 (setq csvfile (open csvPathFile "w"))
	 ;the Header of csv :
	 (write-line "Filename,Plle N,ST" csvfile)
	 
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
		(foreach & Files
			(setq filename (vl-filename-base &)) ; Extract filename without path and extension
			(setq NPlle ""
				  st ""
			)
			(if (is-empty-dwg &)
				(write-line (strcat filename ",Document invalid.") csvfile)
				(progn
					(if (not (is-dwg-open-by-dwl & pcName))
					(progn
						(if (/= (logand (vlax-get-property (vlax-invoke-method FileSystemObject 'getfile &) 'Attributes ) 1) 1)
						(progn
							(setq open-result 
							  (vl-catch-all-apply 'vlax-invoke-method (list dbxDoc 'Open &))
							)
							(if (vl-catch-all-error-p open-result)
							  (write-line (strcat filename ",incompatible version or Document is locked.") csvfile)
							  (progn
								;Start Modify the DWG
								(setq ms (vla-get-modelspace dbxDoc))
								;1 Vlax-For
							    (vlax-for ent ms
								 (if (= (vla-get-objectname ent) "AcDbMText") 
										(if (vl-string-search "Plle N" (vla-get-TextString ent))
										   (setq NPlle (vla-get-TextString ent))
										)
									;)
								  );end if
								  
								  (if (= (vla-get-objectname ent) "AcDbMText") 
										(if (vl-string-search "ST " (vla-get-TextString ent))
										   (setq st (vla-get-TextString ent))
										)
									;)
								  );end if
							    );end 1 vlax-for
							    (princ)
							   
							    (princ)
							    (write-line (strcat filename "," NPlle "," st ) csvfile)
								
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
) ;end DetectNPlle
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(vl-load-com);Load the COM object
(princ)

(princ "\n Lisp Loaded Correctly.")
(princ "\n Let's Start to use DetectNPlle Command :)")
