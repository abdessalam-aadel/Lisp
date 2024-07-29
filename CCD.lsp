;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; CCD : is a small routin for copying content from dwg to others
;;;
;;; Copyright © 2024
;;; https://github.com/abdessalam-aadel/Lisp
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++

(vl-load-com) ; Load COM (Component Object Model) objects

;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;;; Start Main Command CCD
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun c:CCD (/ sourceFile targetFile acapp docs file1 file2 ms1 ms2 originPoint sdi0 filedia0 xloadctl0 demandload0 qaflags0 tilemd0)
  
;Turns off echoing
(Setvar "CMDECHO" 0)

;Set SDI to 0
(if (/= (Getvar "SDI") 0)
	(progn
	(setq sdi0 (Getvar "SDI"))
	(Setvar "SDI" 0)
	)
)
  
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

;Assigns ActiveX Automation interface object for Autodesk AutoCAD to obj
(setq obj (vlax-Get-Acad-Object))

(Prompt "\n Starting process ...\n")
  ;; Create ActiveX Automation interface object for AutoCA
(setq acapp (vlax-create-object "autocad.application")
	docs  (vla-get-documents acapp)
)
  
;; Make AutoCAD visible
(Vlax-put-property acapp "Visible" :vlax-true)
(Vla-put-windowstate acapp acmax)

; Define paths to source and target DWG files
(setq s1 "C:\\dwg\\1.dwg")
(setq s2 "C:\\dwg\\2.dwg")
  
;; Open the target drawing
(Vlax-invoke-method docs "Open" s2)
(Prompt (Strcat "\n Open " s2 ". Please wait ..."))
(setq file2 (Vla-get-activedocument acapp))
(setq ms2 (vla-get-modelspace file2))
  
(Vlax-invoke-method docs "Open" s1)
(Prompt (Strcat "\n Open " s1 ". Please wait ..."))
(setq file1 (Vla-get-activedocument acapp))
(setq ms1 (vla-get-modelspace file1))
  
(vlax-for ent ms1
(if (= (vla-get-ObjectName ent) "AcDbMText")
  (progn
    	(setq insertionPoint (vlax-3d-point (vlax-get ent 'InsertionPoint)))
       (setq textString (vlax-get ent 'TextString))
       (setq height (vlax-get ent 'Height))
       (setq rotation (vlax-get ent 'Rotation))
       (setq width (vla-get-width ent))
       (setq StyleName (vla-get-StyleName ent))
       (setq newText (vla-addMtext ms2 insertionPoint width textString))
       (vla-put-layer newText (vla-get-Layer ent))
       (vla-put-Height newText height)
       (vla-put-Rotation newText rotation)
       (vla-put-StyleName newText StyleName)
    
    )
  )

  (if (= (vla-get-ObjectName ent) "AcDbPolyline")
    (progn
      (prompt "\nbla bla\n")
      ;(setq pts (vla-get-vertices ent)) ;; Get vertices
       (setq vertices '())
      ;; Get the number of vertices
      (setq i 0)
      (while (setq vertex (vlax-get pline 'getPointAt i))
        (setq vertices (append vertices (list (vlax-safearray->list (vlax-get pline 'getPointAt i)))))
        (setq i (1+ i))
      )
      
      ;; Prepare the points for adding a new polyline
      (setq ptArray (vlax-make-safearray vlax-vbDouble (list (* (length vertices) 2) 1)))
      (setq i 0)
      (foreach pt vertices
        (vlax-put ptArray (setq i (1+ i)) (vlax-get pt 'getAt 0))
        (vlax-put ptArray (setq i (1+ i)) (vlax-get pt 'getAt 1))
      )
      ;; Add a new polyline with the extracted vertices
      (setq newPolyline (vla-addLightWeightPolyline ms2 ptArray))
      
      ;(setq polyline (vla-addpolyline ms2 ptArray)) ;; Add polyline
    )
  )
)
  
; Get the TILEMODE Value and store them into tilemode0
(setq tilemd0
  (vlax-variant-value
	(vla-getvariable file2 "TILEMODE" )
  )
)

; Determines whether the Model tab or the most-recently accessed named layout tab is active.
; Set the Model Tab Active
(if (/= tilemd0 1)
 (Vla-SetVariable file2 "TILEMODE" 1)
)

;Turns the grid off
(Vla-SetVariable file2 "GRIDMODE" 0)

;Start Copying
(Prompt "\n Start Copying ...\n")

;; Copy objects from source to target

;; Iterate over each entity in source ModelSpace and copy to target ModelSpace
;;;  (vlax-for ent ms1
;;;    (vla-Add ms2 (vla-copy ent)) ;; Correctly add the entity to the target ModelSpace
;;;  )

   ;; Iterate over each entity in the source ModelSpace and copy to the target ModelSpace
  (vlax-for ent ms1
    (setq entName (vla-get-objectname ent))
    (cond
      ; AcDbText, AcDbMText, AcDbPolyline, AcDbBlockReference
      ((= entName "AcDbLine")
       (setq newLine (vla-addline ms2 (vla-get-startpoint ent) (vla-get-endpoint ent)))
	(vla-put-layer newLine (vla-get-Layer ent))
       )
      
      ((= entName "AcDbCircle")
       (setq newCircle (vla-addcircle ms2 (vla-get-center ent) (vla-get-radius ent)))
	(vla-put-layer newCircle (vla-get-Layer ent))
       )

      ((= entName "AcDbText")
       ;; Handle Text
       (setq insertionPoint (vlax-3d-point (vlax-get ent 'InsertionPoint)))
       (setq textString (vlax-get ent 'TextString))
       (setq height (vlax-get ent 'Height)) ;; Optional: get text height
       (setq rotation (vlax-get ent 'Rotation))
       (setq newText (vla-addtext ms2 textString insertionPoint height))
       (vla-put-layer newText (vla-get-Layer ent))
	)

      ((= entName "AcDbBlockReference")
       ;; Handle BlockReference (Block)
       (setq insertionPoint (vlax-3d-point (vlax-get ent 'InsertionPoint)))
       (setq rotation (vlax-get ent 'Rotation))
       ;; Get block name
       (setq blockName (vlax-get ent 'Name))
       ;; Extract the scale factors from the VLA-OBJECT
      (setq scale-x (vla-get-XScaleFactor ent)) ; X scale factor
      (setq scale-y (vla-get-YScaleFactor ent)) ; Y scale factor
      (setq scale-z (vla-get-ZScaleFactor ent)) ; Z scale factor
      
       ;; Add BlockReference to target drawing
       (setq newBlockRef (vla-insertblock ms2 insertionPoint blockName scale-x scale-y scale-z rotation))
       (vla-put-layer newBlockRef (vla-get-Layer ent))
	)

      ((= entName "AcDbPoint")
       (setq aPoint (vla-get-coordinates ent))
       (setq newPt (vla-addPoint ms2 aPoint))
       (vla-put-layer newPt (vla-get-Layer ent))
	)
    )
  )


  
(princ "\nCopying don!")
  (princ)
  
;Rest the TILEMODE
(if (/= tilemd0 1)
 (Vla-SetVariable file1 "TILEMODE" tilemd0)
)

;Save & Close & Release Object
   ;; Save and close documents
(Prompt (Strcat "\n Save and Close " s2 "\n"))
(vla-save file2)
(vla-close file1)
(vla-close file2)
;; Release COM objects
(vlax-release-object file1)
(vlax-release-object file2)
;close the ActiveX Automation interface object representing AutoCAD
(Vla-quit acapp)
(vlax-release-object docs)
(vlax-release-object acapp)

;Set docs & acapp to null
(setq docs nil
acapp nil
)
(Gc) ;Garbage Collection : explicitly triggers garbage collection
(Gc) ; identify and reclaim memory occupied by objects that are no longer reachable or in use by the program

   ;; Restore system variables
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
  
(princ (Strcat "\n DONE. Processed drawings. !!"))
(setvar "CMDECHO" 1);Turns on echoing
(princ)
) ;end CCD
;;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(princ)

(princ "\n Lisp Loaded Correctly.")
(princ "\n Let's Start to use CCD Command :)")
