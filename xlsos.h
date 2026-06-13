#ifndef XLSOS_H
#define XLSOS_H

#include <stdio.h>
#include <stdint.h>

#ifdef _WIN32
    #define XLSOS_API __declspec(dllexport)
#else
    #define XLSOS_API __attribute__((visibility("default")))
#endif

typedef enum {
    COMPRESS_NONE = 0,
    COMPRESS_ZIP  = 1,
    COMPRESS_7Z   = 2
} CompressMode;

typedef enum {
    FORMAT_XLS  = 0,
    FORMAT_XLSX = 1
} ExcelFormat;

typedef enum {
    CELL_TYPE_NUMBER = 0,
    CELL_TYPE_STRING = 1
} CellType;

typedef struct {
    int32_t row;
    int32_t col;
    CellType type;
    double num_value;
    char* str_value; 
} FastCell;

typedef struct {
    FILE* file_handle;
    ExcelFormat format;
    CompressMode comp_mode;
    char* target_path;  
    char* tmp_dir;      
    int32_t current_row;
} ExcelEngine;

typedef struct {
    ExcelFormat format;
    CompressMode comp_mode;
    const char* file_path;
} EngineConfig;

XLSOS_API ExcelEngine* xlsos_open(const EngineConfig* config, const char* mode);
XLSOS_API int32_t xlsos_write_row(ExcelEngine* engine, const FastCell** cell_ptrs, int32_t* cell_count);
XLSOS_API int32_t xlsos_read_row(ExcelEngine* engine, FastCell*** out_cell_ptrs, int32_t* out_cell_count);
XLSOS_API void xlsos_close(ExcelEngine* engine);

#endif
