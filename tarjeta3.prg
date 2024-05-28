SET EXCLUSIVE OFF
USE t_personal IN 1
SELECT 1
GO TOP

oExcel = CREATEOBJECT("Excel.Application")
oExcel.Visible = .F.
oWorkbook = oExcel.Workbooks.Open("C:\ruta\al\cuentas.xls")
oSheet = oWorkbook.Sheets(1)

LOCAL lnStartTime, lnEndTime
lnStartTime = SECONDS()

DO WHILE NOT EOF()
    FOR i = 2 TO oSheet.UsedRange.Rows.Count
        cNombreExcel = oSheet.Cells(i, 1).Value
        cApellidoExcel = oSheet.Cells(i, 2).Value
        cCedulaExcel = oSheet.Cells(i, 3).Value
        cCuentaExcel = oSheet.Cells(i, 4).Value

        IF t_personal.ps_nombres == cNombreExcel AND t_personal.ps_apellidos == cApellidoExcel AND t_personal.ps_cedula == cCedulaExcel
            REPLACE t_personal.ps_ctabr WITH cCuentaExcel IN t_personal
        ENDIF
    NEXT
    SKIP 1
    IF EOF()
        EXIT
    ENDIF
ENDDO

lnEndTime = SECONDS()
? "Tiempo de procesamiento en VFP: ", lnEndTime - lnStartTime

oWorkbook.Close(.F.)
oExcel.Quit()
RELEASE oExcel
