SET EXCLUSIVE OFF
* Desactiva el modo exclusivo para permitir acceso multiusuario.

USE t_personal IN 1
* Abre la tabla t_personal en el área de trabajo 1.

SELECT 1
* Selecciona el área de trabajo 1.

GO TOP
* Mueve el puntero al primer registro en t_personal.

* Crear un objeto de Excel
oExcel = CREATEOBJECT("Excel.Application")
* Crea un objeto de la aplicación Excel.

oExcel.Visible = .F.
* Hace que la aplicación Excel sea invisible.

* Abrir el libro de Excel
oWorkbook = oExcel.Workbooks.Open("C:\ruta\al\cuentas.xls")
* Abre el archivo Excel especificado.

* Seleccionar la hoja de trabajo
oSheet = oWorkbook.Sheets(1)
* Selecciona la primera hoja del libro Excel.

DO WHILE NOT EOF()
* Inicia un bucle que continúa hasta llegar al final de t_personal.

    * Recorrer las filas del archivo de Excel
    FOR i = 2 TO oSheet.UsedRange.Rows.Count
    * Recorre cada fila de la hoja de Excel (empezando desde la fila 2 para evitar los encabezados).

        cNombreExcel = oSheet.Cells(i, 1).Value
        * Obtiene el valor de la celda de la columna nombres.

        cApellidoExcel = oSheet.Cells(i, 2).Value
        * Obtiene el valor de la celda de la columna apellidos.

        cCedulaExcel = oSheet.Cells(i, 3).Value
        * Obtiene el valor de la celda de la columna cedula.

        cCuentaExcel = oSheet.Cells(i, 4).Value
        * Obtiene el valor de la celda de la columna cuenta.

        * Comprobar si coinciden nombre, apellido y cédula
        IF t_personal.ps_nombres == cNombreExcel AND t_personal.ps_apellidos == cApellidoExcel AND t_personal.ps_cedula == cCedulaExcel
            * Actualizar la tabla t_personal con la cuenta de Excel
            REPLACE t_personal.ps_ctabr WITH cCuentaExcel IN t_personal
        ENDIF

    NEXT

    SKIP 1
    * Mueve al siguiente registro en t_personal.

    IF EOF()
        EXIT
    ENDIF
    * Si se llega al final de t_personal, sale del bucle.

ENDDO

* Cerrar el libro de Excel
oWorkbook.Close(.F.)
* Cierra el libro Excel sin guardar cambios.

oExcel.Quit()
* Cierra la aplicación Excel.

RELEASE oExcel
* Libera el objeto de Excel.

WAIT WINDOW 'proceso terminado'
* Muestra una ventana indicando que el proceso ha terminado.

