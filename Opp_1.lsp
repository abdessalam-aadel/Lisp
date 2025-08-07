;;;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Opp_1 
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
;;; Start function to set the color object to By layer
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun dsx-put-color (obj num / av)
  (setq av (substr (getvar "acadver") 1 2))
  (if (>= av "16")
    ;if AutoCAD 2004...
    (dsx-put-color2004 obj num) 
    ;if any other version...
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
    (progn ;if getting the TrueColor object of 'obj' did not return a vla-error
      (cond
        ( (= 'INT (type num)) ;; if an ACI index integer is passed 
          ;if obj is a table record (i.e. Layer)...
          (if (vl-string-search "Table" (vla-get-ObjectName obj))
            (progn
              (vla-put-ColorMethod oColor acColorMethodByACI) 
              (vla-put-ColorIndex oColor num)
            )
            (if (= num acBylayer) ;; if obj is an entity
              (progn ;if num is to be byLayer 
                (vla-put-ColorMethod oColor acColorMethodByLayer)
                (vla-put-ColorIndex oColor acByLayer)
              )
              (progn ;if num is to be an override 
                (vla-put-ColorMethod oColor acColorMethodByACI)
                (vla-put-ColorIndex oColor num)
              )
            ) 
          )
        )
        ( (and (listp num) (= (length num) 3)) ;an RGB list is passed
          (vla-put-ColorMethod oColor acColorMethodByRGB) ;; set the method 
          ;; set the RGB values
          (vlax-invoke-method oColor 'SetRGB (nth 0 num) (nth 1 num) (nth 2 num)) 
        )
      )
      ;stuff color object back into parent object 
      (vla-put-TrueColor obj oColor)
      ;clean up the memory stack of unused objects
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
;;; Start Main Command Opp_1
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun c:Opp_1 (/   ctf	  	  DwgPath    csvPathFile csvfile     pcName
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
  (setq pcName (getenv "COMPUTERNAME"))
  ;Start Condition
  (cond
    ((setq DwgPath (GetFolder "Selectionnez le dossier qui contient les fichiers autocad :"))
	 ;Set the path of log.csv file
	 (setq csvPathFile (strcat DwgPath "\log.csv"))
	 ;Open the csv file for writing
	 (setq csvfile (open csvPathFile "w"))
	 ;the Header of csv :
	 (write-line "Filename,Opp,x,y,z,Height" csvfile)
	 
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
								(setq ms (vla-get-modelspace dbxDoc))
								
								(vlax-for ent ms
									(if (and (= (vla-get-ObjectName ent) "AcDbMText")
											 (vl-string-search "au profit" (vla-get-TextString ent))
										) 
										(progn
											(setq textStr (vla-get-TextString ent))
											(setq insPt (vlax-safearray->list (vlax-variant-value (vla-get-InsertionPoint ent))))
											(setq HeightTxt (vla-get-Height ent))

											;; Write line to CSV
											(write-line
											  (strcat filename ","
												"\"" textStr "\"," ; text content (quoted to handle commas)
												(rtos (car insPt) 2 3) "," ; X
												(rtos (cadr insPt) 2 3) "," ; Y
												(rtos (caddr insPt) 2 3) "," ; Z
												(rtos HeightTxt 2 3) ; Text height
											  )
											  csvfile
											)
										)
									)
								)
								(princ)
								(setq i (1+ i));iterate the Counter
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
) ;end Opp_1
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(vl-load-com);Load the COM object
(princ)

(princ "\n Lisp Loaded Correctly.")
(princ "\n Let's Start to use Opp_1 Command :)")
