;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; MMO Lisp : is a samll Routine to Set the Color of Titre to Blue and Delete the Big
;;; Polyligne by Layer (RECT), and Delete the Mtext contain Echelle : 1/1000
;;; and Mtext contain the Mappe ex : 33-30-15-A in Multiple DWG Files 
;;;
;;; Copyright © 2024
;;; https://github.com/abdessalam-aadel/PZE
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++

(vl-load-com) ; Load COM (Component Object Model) objects

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

;;; Start Main Command MMO
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun c:MMO (/ acapp      ctf	  demandload0		docs	   DwgPath
			    File       filedia0	  Files	     FilesList	i	   qaflags0
			    sdi0       SubDir	  obj	     tilemd0	xloadctl0
			   )
	
  ;Turns off echoing
  (Setvar "CMDECHO" 0)
  
  ;Set SDI to 0
  (if (/= (Getvar "SDI") 0)
    (progn
      (setq sdi0 (Getvar "SDI"))
      (Setvar "SDI" 0)
    )
  )
  
  ;Assigns ActiveX Automation interface object for Autodesk AutoCAD to obj
  (setq obj (vlax-Get-Acad-Object))
  
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
     (setq Files (mapcar 'strcase Files))
	 
     (Prompt "\n Starting process ...\n")
     (cond
		(Files
		(setq ctf (Length files))
		(setq acapp (vlax-create-object "autocad.application")
			  docs  (vla-get-documents acapp)
		)
	
		(Vlax-put-property acapp "Visible" :vlax-true)
		(Vla-put-windowstate acapp acmax)
	
		;Suppresses display of file navigation dialog boxes.
		;Does not display dialog boxes
		(if (/= (Getvar "FILEDIA") 0)
		  (progn
			(setq filedia0 (Getvar "FILEDIA"))
			(setvar "FILEDIA" 0)
		  )
		)
		
		;Set XLOADCTL = 0 : Turns off demand-loading, the entire drawing is loaded.
		(if (/= (Getvar "XLOADCTL") 0)
		  (progn
			(setq xloadctl0 (Getvar "XLOADCTL"))
			(setvar "XLOADCTL" 0)
		  )
		)
	
		; Turns off demand-loading
		(if (/= (Getvar "DEMANDLOAD") 0)
		  (progn
			(setq demandload0 (Getvar "DEMANDLOAD"))
			(setvar "DEMANDLOAD" 0)
		  )
		)
		
		; Set QFLAGS Quality Assurance (QA) flags to 31 (in binary 11111)
		(if (/= (Getvar "QAFLAGS") 31)
		  (progn
			(setq qaflags0 (Getvar "QAFLAGS"))
			(setvar "QAFLAGS" 31)
		  )
		)
		
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
			  ((Vlax-invoke-method docs "Open" &)
			   (Prompt (Strcat "\n Open " & ". Please wait (" (Itoa (1+ i)) "/" (Itoa ctf) ")..."))
			   (setq file (Vla-get-activedocument acapp))
			   
			   ; Get the TILEMODE Value and store them into tilemode0
			   (setq tilemd0
				  (vlax-variant-value
					(vla-getvariable file "TILEMODE" )
				  )
			   )
			   
			   ; Determines whether the Model tab or the most-recently accessed named layout tab is active.
			   ; Set the Model Tab Active
			   (if (/= tilemd0 1)
				 (Vla-SetVariable file "TILEMODE" 1)
			   )
			   
			   ;Turns the grid off
			   (Vla-SetVariable file "GRIDMODE" 0)
			   
			   ;Start Modify the Mappe
			   (Prompt "\n Start Modify the Mappe ...")
			   (vla-startundomark file)
			   (setq ms (vla-get-modelspace file))

			   ;First Vlax-For
			   (vlax-for ent ms
			     (if (and (= (vla-get-objectname ent) "AcDbPolyline")    
		       		 (= (vla-get-Layer ent) "titres"))
			          (progn
				    (if (not (= (vla-get-Color ent) acBlue))
				      (progn
				       (vla-put-Color ent acBlue)
				       (vla-update ent) ; Update the entity to reflect the color change
				       (princ "\nColor changed to Blue.")
				       (princ)
			       	      )
				    );end if
				  );end progn
		              );end if

			     ;Delete Mtext : Echelle : 1/1000
			     (if (and (= (vla-get-objectname ent) "AcDbMText")    
	       		 	(vl-string-search "Echelle" (vla-get-TextString ent)))
			       	(progn
		       		  (vla-delete ent)
				  (vla-Regen file acActiveViewport)
				)
		              );end if
		      	   );End First Vlax for
			   (princ)

			   ;Second Vlax-For
			   (vlax-for ent ms
			   ;Delete Mtext : 30-30-15 Mappe
			     ;(vlax-invoke-method ent 'Explode)
			     (if (and (= (vla-get-objectname ent) "AcDbMText")    
       		 		  (vl-string-search "-" (vla-get-TextString ent)))
		       		  (progn
		       		   (vla-delete ent)
				   (vla-Regen file acActiveViewport)
				  )
		              );end if
			   
			   );end Second vlax-for
			   (princ)
			   
			   ;Third Vlax-For
			   (vlax-for ent ms
			     (if (and (= (vla-get-objectname ent) "AcDbPolyline")    
		       		 (= (vla-get-Layer ent) "RECT"))
			          (progn
				    (if (> (vla-get-Area ent) 2830170)
				      (progn
				       (vla-delete ent)
				       (vla-Regen file acActiveViewport)
				      )
				    )
				  );end progn
		              );end if
			   );end Third Vlax-For
			   (princ)
			   
			   (vla-endundomark file)
			   ;End Modify the Mappe
			   
			   ;ZoomExtents
			   (vla-Zoomextents (vla-get-application file))
			   
			   ;Rest the TILEMODE
			   (if (/= tilemd0 1)
				 (Vla-SetVariable file "TILEMODE" tilemd0)
			   )
			   
			   ;Start Purge-All
			   (Prompt (Strcat "\n Audit and Purge " & ". Please wait..."))
			   (vla-AuditInfo File T)
			   (vla-purgeall file)
			   (vla-purgeall file)
			   
			   ;Save & Close & Release Object
			   (Prompt (Strcat "\n Save and Close " & "\n"))
			   (vla-save File)
			   (vla-close File)
			   (vlax-release-object File)
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
		;close the ActiveX Automation interface object representing AutoCAD
		(Vla-quit acapp)
		;Release a reference to a COM object
		(Vlax-release-object docs)
		(Vlax-release-object acapp)
		;Set docs & acapp to null
		(setq docs nil
			  acapp nil
		)
		(Gc) ;Garbage Collection : explicitly triggers garbage collection
		(Gc) ; identify and reclaim memory occupied by objects that are no longer reachable or in use by the program
       )
       (T (Prompt "\nNothing files found to purge. "))
     );end cond
    )
    (T (Prompt "\nNothing selected. "))
  );end condition
  
  ;Reset SDI
  (if sdi0
    (progn
      (Setvar "SDI" sdi0)
      (setq sdi0 nil)
    )
  )
  
  ;Reset FILEDIA : file navigation dialog boxes
  (if filedia0
    (progn
      (Setvar "FILEDIA" filedia0)
      (setq filedia0 nil)
    )
  )
  
  ;Reset XLOADCTL : xref demand-loading
  (if xloadctl0
    (progn
      (Setvar "XLOADCTL" xloadctl0)
      (setq xloadctl0 nil)
    )
  )
  
  ;Reset DEMANDLOAD : Components are loaded eagerly at startup.
  (if demandload0
    (progn
      (Setvar "DEMANDLOAD" demandload0)
      (setq demandload0 nil)
    )
  )
  
  ;Reset QAFLAGS Quality Assurance (QA) flags
  (if qaflags0
    (progn
      (Setvar "QAFLAGS" qaflags0)
      (setq qaflags0 nil)
    )
  )
  
  (princ (Strcat "\n DONE. Processed " (Itoa i) " drawings. !!"))
  (setvar "CMDECHO" 1);Turns on echoing
  (princ)
) ;end MMO
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(princ)

(princ "\n Lisp Loaded Correctly.")
(princ "\n Let's Start to use MMO Command :)")
