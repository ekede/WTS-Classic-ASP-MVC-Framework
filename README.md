### WTS Classic ASP Framework with VBScript based on Object ###

Code: UTF-8 with BOM

Program: request => route => module/control/action (model,view,language) => response

Start: inc/wts.start() => inc/module/---/control/start/site.start() => ...

Struct:

    index.asp        Single entry point, IIS404,405 point here

    inc/             Program Folder
    inc/config.asp   Global configure File
    inc/wts.asp      Core Framework 
    inc/class/       Asp Libray
    inc/module/      MVCL Program

    data/            Data Folder
    data/cache/      Cache Folder
    data/db/         Database Folder
    data/log/        Log Folder
    data/pic/        IMAGE Folder
    data/static/     Static Folder: css,js,icon...

    app/             Rewrite "inc/" Folder

Note:

    Program and Data path can be changed by config.asp
    Class libraries and program files based on object, loading file by loader object