(defun c:EP ( / acadApp doc ms ss i ent obj newDoc fname path )

  (vl-load-com)

  ;; -- Find nearest text to given polyline center --
  (defun find-nearest-text (x y / ss numEnt i ent obj textPos dist minDist text)
    (setq text "" minDist 1.0e+30)
    (setq ss (ssget "_X" '((0 . "TEXT,MTEXT"))))
    (if ss
      (progn
        (setq numEnt (sslength ss) i 0)
        (repeat numEnt
          (setq ent (ssname ss i))
          (if (and ent (setq obj (vlax-ename->vla-object ent)))
            (progn
              (setq textPos (vlax-get obj 'InsertionPoint))
              (setq dist (distance (list x y 0.0) textPos))
              (if (< dist minDist)
                (progn
                  (setq minDist dist)
                  (setq text (vla-get-TextString obj))
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

  ;; -- Get name of polyline based on center point --
  (defun get-polyline-label (obj / coords numVerts cx cy i text)
    (if (vlax-property-available-p obj 'Coordinates)
      (progn
        (setq coords (vlax-get obj 'Coordinates))
        (setq numVerts (/ (length coords) 2))
        (setq cx 0.0 cy 0.0 i 0)
        (repeat numVerts
          (setq cx (+ cx (nth (* i 2) coords)))
          (setq cy (+ cy (nth (1+ (* i 2)) coords)))
          (setq i (1+ i))
        )
        (setq cx (/ cx numVerts))
        (setq cy (/ cy numVerts))
        (setq text (find-nearest-text cx cy))

        ;; Fallback if no text found
        (if (or (not text) (= text "")) (setq text "Polyline"))

        ;; Sanitize filename
        (foreach ch '("\\" "/" ":" "*" "?" "\"" "<" ">" "|")
          (setq text (vl-string-subst "_" ch text))
        )
        text
      )
      "Polyline"
    )
  )

  ;; -- Main Routine --
  (setq acadApp (vlax-get-acad-object))
  (setq doc (vla-get-ActiveDocument acadApp))
  (setq ms (vla-get-ModelSpace doc))
  (setq ss (ssget "X" '((0 . "LWPOLYLINE,POLYLINE"))))

  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq obj (vlax-ename->vla-object ent))

        ;; Create new drawing
        (setq newDoc (vla-Add (vla-get-Documents acadApp)))
        (vla-Activate newDoc)

        ;; Copy polyline into new drawing
        (vla-CopyObjects
          doc
          (vlax-make-variant
            (vlax-safearray-fill
              (vlax-make-safearray vlax-vbObject '(0 . 0))
              (list obj)))
          (vla-get-ModelSpace newDoc)
        )

        ;; Generate safe file name
        (setq fname (get-polyline-label obj))
        (setq path (strcat (getvar "DWGPREFIX") fname ".dwg"))
		
		(vla-Activate doc) ; Switch back to the original drawing

        ;; Save and close new drawing
        (vla-SaveAs newDoc path)
        (vla-Close newDoc :vlax-false)

        (setq i (1+ i))
      )
      (princ (strcat "\nExported " (itoa i) " polylines."))
    )
    (prompt "\nNo polylines found.")
  )
  (princ)
)
