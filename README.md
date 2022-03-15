ファイルの配置

projectFolder
│   CloseWindoww.py
|   export.py
|   import.py
|   vbac.wsf
+---bin
|      book.xlsm
\---src
    \---book.xlsm
        Module1.bas
        Module2.bas
        Module3.bas


・export.py
binファイルに入っているxlsmファイルのbas、dcmファイルをエクスポートする
エクスポートはブックを開いていても可能

・import.py
bas、dcmファイルをbookファイルにインポートする
インポートはブックを閉じないと実行できない為、ブックを閉じた状態で実行するとブックを閉じるよう要求される

・CloseWindow.py
ウィンドウハンドルで指定されたウィンドウを閉じるファイル
このファイルではbinファイルに入っているbookを閉じるようにしている
