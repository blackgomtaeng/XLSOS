#include "xlsos.h"
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#ifdef _WIN32
    #include <direct.h>
    #define MAKE_DIR(dir) _mkdir(dir)
    #define RM_DIR(dir) { size_t s = strlen(dir) + 32; char* c = (char*)malloc(s); snprintf(c, s, "rmdir /s /q \"%s\" > nul 2>&1", dir); system(c); free(c); }
#else
    #include <sys/stat.h>
    #include <sys/types.h>
    #define MAKE_DIR(dir) mkdir(dir, 0755)
    #define RM_DIR(dir) { size_t s = strlen(dir) + 32; char* c = (char*)malloc(s); snprintf(c, s, "rm -rf \"%s\" > /dev/null 2>&1", dir); system(c); free(c); }
#endif

typedef struct {
    uint16_t code;
    uint16_t length;
} BiffHeader;

ExcelEngine* xlsos_open(const EngineConfig* config, const char* mode) {
    if (!config || !mode || !config->file_path) return NULL;

    ExcelEngine* engine = (ExcelEngine*)malloc(sizeof(ExcelEngine));
    if (!engine) return NULL;

    engine->format = config->format;
    engine->comp_mode = config->comp_mode;
    engine->current_row = 0;
    
    size_t path_len = strlen(config->file_path);
    engine->target_path = (char*)malloc(path_len + 1);
    strcpy(engine->target_path, config->file_path);

    size_t tmp_alloc_size = path_len + 5;
    engine->tmp_dir = (char*)malloc(tmp_alloc_size);
    snprintf(engine->tmp_dir, tmp_alloc_size, "%s_tmp", config->file_path);

    if (strcmp(mode, "w") == 0 || strcmp(mode, "wb") == 0) {
        if (engine->format == FORMAT_XLS) {
            engine->file_handle = fopen(engine->target_path, "wb");
            if (!engine->file_handle) { free(engine->target_path); free(engine->tmp_dir); free(engine); return NULL; }
            
            BiffHeader bof = { 0x0009, 0x0004 };
            uint16_t bof_version = 0x0002;
            fwrite(&bof, sizeof(BiffHeader), 1, engine->file_handle);
            fwrite(&bof_version, sizeof(uint16_t), 1, engine->file_handle);
        } 
        else {
            RM_DIR(engine->tmp_dir);
            MAKE_DIR(engine->tmp_dir);
            
            size_t buf_size = strlen(engine->tmp_dir) + 32;
            char* path_buffer = (char*)malloc(buf_size);
            
            snprintf(path_buffer, buf_size, "%s/xl", engine->tmp_dir); MAKE_DIR(path_buffer);
            snprintf(path_buffer, buf_size, "%s/xl/worksheets", engine->tmp_dir); MAKE_DIR(path_buffer);
            snprintf(path_buffer, buf_size, "%s/xl/worksheets/sheet1.xml", engine->tmp_dir);
            
            engine->file_handle = fopen(path_buffer, "w");
            free(path_buffer);
            if (!engine->file_handle) { free(engine->target_path); free(engine->tmp_dir); free(engine); return NULL; }

            fprintf(engine->file_handle, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            fprintf(engine->file_handle, "<worksheet xmlns=\"http://openxmlformats.org\">\n");
            fprintf(engine->file_handle, "<sheetData>\n");
        }
    }
    else if (strcmp(mode, "r") == 0 || strcmp(mode, "rb") == 0) {
        if (engine->format == FORMAT_XLS) {
            engine->file_handle = fopen(engine->target_path, "rb");
            if (!engine->file_handle) { free(engine->target_path); free(engine->tmp_dir); free(engine); return NULL; }
            BiffHeader bof;
            if (fread(&bof, sizeof(BiffHeader), 1, engine->file_handle) == 1) fseek(engine->file_handle, bof.length, SEEK_CUR);
        } 
        else {
            RM_DIR(engine->tmp_dir);
            MAKE_DIR(engine->tmp_dir);

            size_t cmd_size = strlen(engine->target_path) + strlen(engine->tmp_dir) + 128;
            char* cmd = (char*)malloc(cmd_size);
            if (engine->comp_mode == COMPRESS_ZIP) {
#ifdef _WIN32
                snprintf(cmd, cmd_size, "powershell Expand-Archive -Path \"%s\" -DestinationPath \"%s\" -Force > nul", engine->target_path, engine->tmp_dir);
#else
                snprintf(cmd, cmd_size, "unzip -q \"%s\" -d \"%s\" > /dev/null 2>&1", engine->target_path, engine->tmp_dir);
#endif
                system(cmd);
            }
            free(cmd);

            size_t buf_size = strlen(engine->tmp_dir) + 64;
            char* path_buffer = (char*)malloc(buf_size);
            snprintf(path_buffer, buf_size, "%s/xl/worksheets/sheet1.xml", engine->tmp_dir);
            engine->file_handle = fopen(path_buffer, "r");
            free(path_buffer);
            if (!engine->file_handle) { free(engine->target_path); free(engine->tmp_dir); free(engine); return NULL; }
        }
    }

    return engine;
}

int32_t xlsos_write_row(ExcelEngine* engine, const FastCell** cell_ptrs, int32_t* cell_count) {
    if (!engine || !engine->file_handle || !cell_ptrs || !cell_count || *cell_count <= 0) return -1;
    engine->current_row++;

    const FastCell** current_pp = cell_ptrs;

    if (engine->format == FORMAT_XLS) {
        for (int32_t i = 0; i < *cell_count; i++) {
            const FastCell* cell = *current_pp;
            if (cell->type == CELL_TYPE_NUMBER) {
                BiffHeader num_hdr = { 0x0003, 11 };
                uint16_t r = (uint16_t)(engine->current_row - 1);
                uint16_t c = (uint16_t)cell->col;
                uint16_t attr = 0x0000;
                double val = cell->num_value;

                fwrite(&num_hdr, sizeof(BiffHeader), 1, engine->file_handle);
                fwrite(&r, sizeof(uint16_t), 1, engine->file_handle);
                fwrite(&c, sizeof(uint16_t), 1, engine->file_handle);
                fwrite(&attr, sizeof(uint16_t), 1, engine->file_handle);
                fwrite(&val, sizeof(double), 1, engine->file_handle);
            } 
            else if (cell->type == CELL_TYPE_STRING) {
                uint8_t str_len = (uint8_t)strlen(cell->str_value);
                BiffHeader str_hdr = { 0x0004, (uint16_t)(7 + str_len) };
                uint16_t r = (uint16_t)(engine->current_row - 1);
                uint16_t c = (uint16_t)cell->col;
                uint16_t attr = 0x0000;

                fwrite(&str_hdr, sizeof(BiffHeader), 1, engine->file_handle);
                fwrite(&r, sizeof(uint16_t), 1, engine->file_handle);
                fwrite(&c, sizeof(uint16_t), 1, engine->file_handle);
                fwrite(&attr, sizeof(uint16_t), 1, engine->file_handle);
                fwrite(&str_len, sizeof(uint8_t), 1, engine->file_handle);
                fwrite(cell->str_value, sizeof(char), str_len, engine->file_handle);
            }
            current_pp++; 
        }
    } 
    else {
        fprintf(engine->file_handle, "<row r=\"%d\">\n", engine->current_row);
        for (int32_t i = 0; i < *cell_count; i++) {
            const FastCell* cell = *current_pp;
            char col_letter = 'A' + (cell->col % 26);
            if (cell->type == CELL_TYPE_NUMBER) fprintf(engine->file_handle, "<c r=\"%c%d\"><v>%.6f</v></c>\n", col_letter, engine->current_row, cell->num_value);
            else fprintf(engine->file_handle, "<c r=\"%c%d\" t=\"inlineStr\"><is><t>%s</t></is></c>\n", col_letter, engine->current_row, cell->str_value);
            current_pp++;
        }
        fprintf(engine->file_handle, "</row>\n");
    }
    return engine->current_row;
}

int32_t xlsos_read_row(ExcelEngine* engine, FastCell*** out_cell_ptrs, int32_t* out_cell_count) {
    if (!engine || !engine->file_handle || !out_cell_ptrs || !out_cell_count) return -1;

    *out_cell_count = 0;
    *out_cell_ptrs = NULL;
    
    int32_t capacity = 4; 
    FastCell** ptr_array = (FastCell**)malloc(sizeof(FastCell*) * capacity);

    if (engine->format == FORMAT_XLS) {
        BiffHeader hdr;
        while (fread(&hdr, sizeof(BiffHeader), 1, engine->file_handle) == 1) {
            if (hdr.code == 0x0003 || hdr.code == 0x0004) {
                if (*out_cell_count >= capacity) {
                    capacity *= 2;
                    ptr_array = (FastCell**)realloc(ptr_array, sizeof(FastCell*) * capacity);
                }
                
                FastCell* target_cell = (FastCell*)malloc(sizeof(FastCell));
                uint16_t r, c, attr;
                fread(&r, sizeof(uint16_t), 1, engine->file_handle);
                fread(&c, sizeof(uint16_t), 1, engine->file_handle);
                fread(&attr, sizeof(uint16_t), 1, engine->file_handle);

                target_cell->row = r;
                target_cell->col = c;

                if (hdr.code == 0x0003) {
                    fread(&(target_cell->num_value), sizeof(double), 1, engine->file_handle);
                    target_cell->type = CELL_TYPE_NUMBER;
                    target_cell->str_value = NULL;
                } else {
                    uint8_t str_len;
                    fread(&str_len, sizeof(uint8_t), 1, engine->file_handle);
                    target_cell->str_value = (char*)malloc(str_len + 1);
                    fread(target_cell->str_value, sizeof(char), str_len, engine->file_handle);
                    target_cell->str_value[str_len] = '\0';
                    target_cell->type = CELL_TYPE_STRING;
                }
                
                *(ptr_array + (*out_cell_count)) = target_cell;
                (*out_cell_count)++;
                
                long curr_pos = ftell(engine->file_handle);
                BiffHeader next_hdr;
                if (fread(&next_hdr, sizeof(BiffHeader), 1, engine->file_handle) == 1) {
                    uint16_t next_r;
                    fread(&next_r, sizeof(uint16_t), 1, engine->file_handle);
                    fseek(engine->file_handle, curr_pos, SEEK_SET);
                    if (next_r != r) break;
                }
            } 
            else if (hdr.code == 0x000A) break;
            else fseek(engine->file_handle, hdr.length, SEEK_CUR);
        }
    } 
    else {
        size_t line_cap = 512;
        char* line = (char*)malloc(line_cap);
        size_t line_len = 0;
        int ch;
        
        while ((ch = fgetc(engine->file_handle)) != EOF) {
            if (ch == '\n') {
                line[line_len] = '\0';
                if (strstr(line, "<row") != NULL) {
                    char* cell_ptr = line;
                    while ((cell_ptr = strstr(cell_ptr, "<c ")) != NULL) {
                        if (*out_cell_count >= capacity) {
                            capacity *= 2;
                            ptr_array = (FastCell**)realloc(ptr_array, sizeof(FastCell*) * capacity);
                        }
                        
                        FastCell* target_cell = (FastCell*)malloc(sizeof(FastCell));
                        target_cell->str_value = NULL;

                        char* r_attr = strstr(cell_ptr, "r=\"");
                        if (r_attr) {
                            r_attr += 3;
                            char* r_end = strchr(r_attr, '"');
                            if (r_end) {
                                *r_end = '\0';
                                char col_letter = r_attr[0];
                                int row_num = atoi(r_attr + 1);
                                target_cell->col = toupper(col_letter) - 'A';
                                target_cell->row = row_num;
                            }
                        }

                        if (strstr(cell_ptr, "t=\"inlineStr\"") != NULL) {
                            target_cell->type = CELL_TYPE_STRING;
                            char* t_start = strstr(cell_ptr, "<t>");
                            char* t_end = strstr(cell_ptr, "</t>");
                            if (t_start && t_end && t_end > t_start) {
                                t_start += 3;
                                size_t str_len = t_end - t_start;
                                target_cell->str_value = (char*)malloc(str_len + 1);
                                strncpy(target_cell->str_value, t_start, str_len);
                                target_cell->str_value[str_len] = '\0';
                            }
                        } else {
                            target_cell->type = CELL_TYPE_NUMBER;
                            char* v_start = strstr(cell_ptr, "<v>");
                            char* v_end = strstr(cell_ptr, "</v>");
                            if (v_start && v_end && v_end > v_start) {
                                v_start += 3;
                                char num_buf[64];
                                size_t num_len = v_end - v_start;
                                strncpy(num_buf, v_start, num_len);
                                num_buf[num_len] = '\0';
                                target_cell->num_value = atof(num_buf);
                            }
                        }

                        *(ptr_array + (*out_cell_count)) = target_cell;
                        (*out_cell_count)++;
                        cell_ptr++;
                    }
                }
                line_len = 0;
            } else {
                if (line_len < line_cap - 1) line[line_len++] = (char)ch;
            }
        }
        free(line);
    }
    *out_cell_ptrs = ptr_array;
    return *out_cell_count;
}

void xlsos_close(ExcelEngine* engine) {
    if (!engine) return;

    if (engine->file_handle) {
        if (engine->format == FORMAT_XLSX) {
            fprintf(engine->file_handle, "</sheetData>\n</worksheet>\n");
            fclose(engine->file_handle);

            size_t cmd_size = strlen(engine->target_path) + strlen(engine->tmp_dir) + 128;
            char* cmd = (char*)malloc(cmd_size);
            if (engine->comp_mode == COMPRESS_ZIP) {
#ifdef _WIN32
                snprintf(cmd, cmd_size, "powershell Compress-Archive -Path \"%s\" -DestinationPath \"%s\" -Force > nul", engine->tmp_dir, engine->target_path);
#else
                snprintf(cmd, cmd_size, "zip -qr \"%s\" \"%s\" > /dev/null 2>&1", engine->target_path, engine->tmp_dir);
#endif
                system(cmd);
            }
            free(cmd);
            RM_DIR(engine->tmp_dir);
        } else {
            fclose(engine->file_handle);
        }
    }
    free(engine->target_path);
    free(engine->tmp_dir);
    free(engine);
}
