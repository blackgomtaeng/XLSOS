#include "xlsos.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

typedef struct {
    int32_t uid;
    double score;
    char* name; 
} UserRecord;

int main(int argc, char* argv[]) {
    EngineConfig run_config;
    run_config.format = FORMAT_XLSX; 
    run_config.comp_mode = COMPRESS_ZIP;
    run_config.file_path = "multi_pointer_test.xlsx";

    int32_t dynamic_sample_count = 3;
    UserRecord** record_ptrs = (UserRecord**)malloc(sizeof(UserRecord*) * dynamic_sample_count);

    UserRecord** current_rec = record_ptrs;
    
    *current_rec = (UserRecord*)malloc(sizeof(UserRecord));
    (*current_rec)->uid = 5001; (*current_rec)->score = 99.87;
    (*current_rec)->name = (char*)malloc(strlen("Dynamic_Engine_Alpha_Node") + 1);
    strcpy((*current_rec)->name, "Dynamic_Engine_Alpha_Node");
    current_rec++;

    *current_rec = (UserRecord*)malloc(sizeof(UserRecord));
    (*current_rec)->uid = 5002; (*current_rec)->score = 84.13;
    (*current_rec)->name = (char*)malloc(strlen("Beta_Core") + 1);
    strcpy((*current_rec)->name, "Beta_Core");
    current_rec++;

    *current_rec = (UserRecord*)malloc(sizeof(UserRecord));
    (*current_rec)->uid = 5003; (*current_rec)->score = 91.50;
    (*current_rec)->name = (char*)malloc(strlen("Gamma_Ultra_Sub_System_Node") + 1);
    strcpy((*current_rec)->name, "Gamma_Ultra_Sub_System_Node");

    ExcelEngine* write_engine = xlsos_open(&run_config, "wb");
    if (write_engine) {
        int32_t cells_per_row = 3;
        FastCell** packet_ptrs = (FastCell**)malloc(sizeof(FastCell*) * cells_per_row);

        current_rec = record_ptrs; 
        for(int32_t i = 0; i < dynamic_sample_count; i++) {
            FastCell** p_cursor = packet_ptrs;

            *p_cursor = (FastCell*)malloc(sizeof(FastCell));
            (*p_cursor)->col = 0; (*p_cursor)->type = CELL_TYPE_NUMBER; (*p_cursor)->num_value = (double)(*current_rec)->uid;
            p_cursor++;

            *p_cursor = (FastCell*)malloc(sizeof(FastCell));
            (*p_cursor)->col = 1; (*p_cursor)->type = CELL_TYPE_STRING; 
            (*p_cursor)->str_value = (char*)malloc(strlen((*current_rec)->name) + 1);
            strcpy((*p_cursor)->str_value, (*current_rec)->name);
            p_cursor++;

            *p_cursor = (FastCell*)malloc(sizeof(FastCell));
            (*p_cursor)->col = 2; (*p_cursor)->type = CELL_TYPE_NUMBER; (*p_cursor)->num_value = (*current_rec)->score;

            xlsos_write_row(write_engine, (const FastCell**)packet_ptrs, &cells_per_row);

            p_cursor = packet_ptrs;
            for(int32_t c = 0; c < cells_per_row; c++) {
                if ((*p_cursor)->type == CELL_TYPE_STRING) free((*p_cursor)->str_value);
                free(*p_cursor);
                p_cursor++;
            }
            current_rec++;
        }
        free(packet_ptrs);
        xlsos_close(write_engine);
    }

    ExcelEngine* read_engine = xlsos_open(&run_config, "rb");
    if (read_engine) {
        FastCell** incoming_row_ptrs = NULL; 
        int32_t retrieved_cell_count = 0;

        while (xlsos_read_row(read_engine, &incoming_row_ptrs, &retrieved_cell_count) > 0) {
            FastCell** cell_pp = incoming_row_ptrs; 
            for(int32_t j = 0; j < retrieved_cell_count; j++) {
                FastCell* cell = *cell_pp;
                if (cell->type == CELL_TYPE_NUMBER) printf("Num: %.2f\t", cell->num_value);
                else {
                    printf("Str: %s\t", cell->str_value);
                    free(cell->str_value); 
                }
                free(cell); 
                cell_pp++;  
            }
            printf("\n");
            free(incoming_row_ptrs); 
        }
        xlsos_close(read_engine);
    }

    current_rec = record_ptrs;
    for(int32_t i = 0; i < dynamic_sample_count; i++) {
        free((*current_rec)->name);
        free(*current_rec);
        current_rec++;
    }
    free(record_ptrs);

    return 0;
}