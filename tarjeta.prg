SET EXCLUSIVE OFF

USE t_personal IN 1
**USE libro2 IN 2

SELECT 1
GO TOP

DO WHILE NOT EOF()
** SELECT libro2
UPDATE libro2  set libro2.cedula = t_personal.ps_cedula WHERE t_personal.ps_nombres = libro2.nombres AND t_personal.ps_apellidos = libro2.apellidos

SELECT t_personal 
SKIP 1
IF EOF()
   EXIT
ENDIF

ENDDO
WAIT WINDOW 'proceso terminado'   

