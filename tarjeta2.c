#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <libxls/xls.h>
#include <sqlite3.h>
#include <time.h>

void update_database(const char *db_filename, const char *excel_filename) {
    sqlite3 *db;
    sqlite3_stmt *stmt;
    int rc;

    rc = sqlite3_open(db_filename, &db);
    if (rc) {
        fprintf(stderr, "Can't open database: %s\n", sqlite3_errmsg(db));
        return;
    }

    const char *sql = "UPDATE t_personal SET ps_ctabr = ? WHERE ps_nombres = ? AND ps_apellidos = ? AND ps_cedula = ?";
    rc = sqlite3_prepare_v2(db, sql, -1, &stmt, NULL);
    if (rc != SQLITE_OK) {
        fprintf(stderr, "Failed to prepare statement: %s\n", sqlite3_errmsg(db));
        sqlite3_close(db);
        return;
    }

    xlsWorkBook *workbook = xls_open(excel_filename, "UTF-8");
    if (!workbook) {
        fprintf(stderr, "Failed to open Excel file.\n");
        sqlite3_finalize(stmt);
        sqlite3_close(db);
        return;
    }

    xlsWorkSheet *sheet = xls_getWorkSheet(workbook, 0);
    xls_parseWorkSheet(sheet);

    clock_t start = clock();

    for (int i = 1; i <= sheet->rows.lastrow; i++) {
        xlsRow *row = &sheet->rows.row[i];
        char *nombre = (char *)row->cells.cell[0].str;
        char *apellido = (char *)row->cells.cell[1].str;
        char *cedula = (char *)row->cells.cell[2].str;
        char *cuenta = (char *)row->cells.cell[3].str;

        sqlite3_bind_text(stmt, 1, cuenta, -1, SQLITE_STATIC);
        sqlite3_bind_text(stmt, 2, nombre, -1, SQLITE_STATIC);
        sqlite3_bind_text(stmt, 3, apellido, -1, SQLITE_STATIC);
        sqlite3_bind_text(stmt, 4, cedula, -1, SQLITE_STATIC);

        rc = sqlite3_step(stmt);
        if (rc != SQLITE_DONE) {
            fprintf(stderr, "Update failed: %s\n", sqlite3_errmsg(db));
        }

        sqlite3_reset(stmt);
    }

    clock_t end = clock();
    double time_taken = ((double)(end - start)) / CLOCKS_PER_SEC;
    printf("Tiempo de procesamiento en C: %f segundos\n", time_taken);

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
