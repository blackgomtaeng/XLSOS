# XLSOS

## 01 Windows 빌드 명령어
```
gcc -O2 -shared xlsos.c -o xlsos.dll -Wl,--out-implib,xlsos.a
```

## 02 Windows 실행 파일
```
gcc -O2 main.c -L. -lxlsos -o xlsos_tool.exe
```


## Linux/Unix 빌드 명령어
```
x86_64-linux-gnu-gcc -O2 -fPIC -shared xlsos.c -o xlsos.so
```

## 02 Linux/Unix 실행 파일
```
gcc -O2 main.c -L. -lxlsos -o xlsos_tool -Wl,-rpath,.
```