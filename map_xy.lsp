(defun c:map_xy ()
  (vl-load-com)

  ;; Function to extract vertices from a single polyline and write to file
  (defun extract-polyline-vertices (obj file)
    (if (and (vlax-property-available-p obj 'Coordinates) (eq (vla-get-ObjectName obj) "AcDbPolyline"))
      (progn
        (setq coords (vlax-get obj 'Coordinates)) ; Get the coordinates array
        
        ;; Calculate the center point of the polyline
        (setq numVerts (/ (length coords) 2))
        (setq centerX 0.0)
        (setq centerY 0.0)
        (setq i 0)
        (repeat numVerts
          (setq centerX (+ centerX (nth (* i 2) coords)))
          (setq centerY (+ centerY (nth (1+ (* i 2)) coords)))
          (setq i (1+ i))
        )
        (setq centerX (/ centerX numVerts))
        (setq centerY (/ centerY numVerts))
        
        ;; Find text near the center of the polyline
        (setq text (find-nearest-text centerX centerY))
        
        ;; Write polyline data to file
        ;;(write-line (strcat "Polyline ID: " (itoa (vla-get-ObjectID obj))) file)
        ;;(write-line (strcat "Center: X=" (rtos centerX 2 4) ", Y=" (rtos centerY 2 4)) file)
        (write-line (strcat "" text ":") file)
        ;;(write-line "Vertex Coordinates:" file)
        
        ;; Write vertex coordinates
        (setq i 3)
        (repeat numVerts
          (setq x (nth (* i 2) coords))
          (setq y (nth (1+ (* i 2)) coords))
	  (if (= i 3)
  	      (write-line (strcat (rtos x 2 4) ";" (rtos y 2 4) ":")  file)
	  )
	  (if (= i 1)
  	      (write-line (strcat (rtos x 2 4) ";" (rtos y 2 4)) file)
	  )
          
          (setq i (1- i))
        )
      )
      (write-line "Selected object is not a polyline." file)
    )
  )

  ;; Function to find the nearest text to a given point
  (defun find-nearest-text (x y)
    (setq text "")
    (setq minDist 1.0e+30) ; Large number to represent "infinity"
    (setq ss (ssget "_X" '((0 . "TEXT, MTEXT")))) ; Select all text entities
    (if ss
      (progn
        (setq numEnt (sslength ss))
        (setq i 0)
        (repeat numEnt
          (setq ent (ssname ss i))
          (if ent ; Check if the entity is valid
            (progn
              (setq obj (vlax-ename->vla-object ent))
              (setq textPos (vlax-get obj 'InsertionPoint))
              (setq dist (distance (list x y) textPos))
              (if (< dist minDist)
                (progn
                  (setq minDist dist)
                  (setq text (if (equal (vla-get-ObjectName obj) "AcDbText")
                                 (vla-get-TextString obj)
                                 (vla-get-TextString obj))) ; Handle TEXT and MTEXT
                )
              )
            )
          )
          (setq i (1+ i))
        )
      )
    )
    text
  )

  ;; Main function to export all polyline data to a file
  (defun process-all-polylines (file)
    (setq acadApp (vlax-get-acad-object))
    (setq doc (vla-get-ActiveDocument acadApp))
    (setq ms (vla-get-ModelSpace doc))
    
    (vlax-for ent ms
      (if (and (eq (vla-get-ObjectName ent) "AcDbPolyline") ; Check if the entity is a polyline
               (vlax-property-available-p ent 'Coordinates))
        (extract-polyline-vertices ent file)
      )
    )
  )

  ;; Ask the user to specify the file path
  (setq filepath (getfiled "Select or Enter Output File" "" "txt" 1))
  
  (if filepath
    (progn
      ;; Open the file for writing
      (setq file (open filepath "w"))
      
      ;; Start processing all polylines
      (process-all-polylines file)
      
      ;; Close the file
      (close file)
      
      (princ (strcat "\nData has been written to: " filepath))
    )
    (princ "\nNo file specified.")
  )
  
  (princ)
)
