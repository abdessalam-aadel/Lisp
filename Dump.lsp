;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;; prints out all the properties and methods of a given ActiveX/COM object (obj)
;;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
(defun c:dump ()
	(setq e (entsel)); pick a entity
	(setq obj (vlax-ename->vla-object (car e)))
	(princ (vlax-dump-object obj T))
)