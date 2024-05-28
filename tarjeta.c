#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <libxls/xls.h>
#include <sqlite3.h>

// Función para leer la base de datos y actualizar registros
void update_database(const char *db_filename, const char *excel_filename) {
    sqlite3 *db;
    sqlite3_stmt *stmt;
    int rc;

    // Abrir la base de datos
    rc = sqlite3_open(db_filename, &db);
    if (rc) {
        fprintf(stderr, "Can't open database: %s\n", sqlite3_errmsg(db));
        return;
    }

    // Preparar la consulta de actualización
    const char *sql = "UPDATE t_personal SET ps_ctabr = ? WHERE ps_nombres = ? AND ps_apellidos = ? AND ps_cedula = ?";
    rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
    if (rc != SQLITE_OK) {
        fprintf(stderr, "Failed to prepare statement: %s\n", sqlite3_errmsg(db));
        sqlite3_close(db);
        return;
    }

    // Abrir el archivo de Excel
    xlsWorkBook *workbook = xls_open(excel_filename, "UTF-8");
    if (!workbook) {
        fprintf(stderr, "Failed to open Excel file.\n");
        sqlite3_finalize(stmt);
        sqlite3_close(db);
        return;
    }

    xlsWorkSheet *sheet = xls_getWorkSheet(workbook, 0);
    xls_parseWorkSheet(sheet);

    // Recorrer las filas del archivo de Excel
    for (int i = 1; i <= sheet->rows.lastrow; i++) {
        xlsRow *row = &sheet->rows.row[i];
        char *nombre = (char *)row->cells.cell[0].str;
        char *apellido = (char *)row->cells.cell[1].str;
        char *cedula = (char *)row->cells.cell[2].str;
        char *cuenta = (char *)row->cells.cell[3].str;

        // Enlazar los valores de la fila a la consulta SQL
        sqlite3_bind_text(stmt, 1, cuenta, -1, SQLITE_STATIC);
        sqlite3_bind_text(stmt, 2, nombre, -1, SQLITE_STATIC);
        sqlite3_bind_text(stmt, 3, apellido, -1, SQLITE_STATIC);
        sqlite3_bind_text(stmt, 4, cedula, -1, SQLITE_STATIC);

        // Ejecutar la consulta
        rc = sqlite3_step(stmt);
        if (rc != SQLITE_DONE) {
            fprintf(stderr, "Update failed: %s\n", sqlite3_errmsg(db));
        }

        // Reiniciar la consulta para la siguiente iteración
        sqlite3_reset(stmt);
    }

    // Liberar los recursos
    xls_close_WS(sheet);
    xls_close(workbook);
    sqlite3_finalize(stmt);
    sqlite3_close(db);
}

int main() {
    const char *db_filename = "t_personal.db";
    const char *excel_filename = "C:\\ruta\\al\\cuentas.xls";

    update_database(db_filename, excel_filename);

    printf("Proceso terminado\n");
    return 0;
}
