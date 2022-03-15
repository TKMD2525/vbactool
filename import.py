import os
import subprocess
import time
from CloseWindow import Strategy

def importvba():
    # NOTE:bookを閉じてインポートを実行するようにする
    closing = Strategy()

    # 実行する前に指定したウィンドウが開いているかを確認する
    # someHandle: desktopWinHndかWinHndのどっちかが入っている変数
    someHandle = Strategy.getWindowhandle(closing.title_name, closing.class_name)
    desktopWinHnd = closing.desktopWinHnd

    # ウィンドウが開いていればFalse、開いていればTrue
    openingwindow = someHandle != desktopWinHnd

    # ファイルが開かれている時のみウィンドウを閉じる処理を行う
    if openingwindow == True:
        #ブックを閉じる
        closed = closing()

        #指定したウィンドウが閉じるまで繰り返す
        while not closed:
            time.sleep(0.5)
            #closing(closeメソッド)は、ウィンドウが閉じるとTrueを返す
            # その間は何も返さない → Noneである
            if closing() == None:
                closed = False
            elif closing() == True:
                closed = True

        print("指定されたウィンドウを閉じたことを確認しました。")

    #インポート
    print("\nインポートを実行します。\n")
    subprocess.call('cscript //nologo vbac.wsf combine', shell=True)

def main():
    # 自分の相対パス取得
    path = os.path.dirname(__file__)

    # 自分のファイルパスにcd
    os.chdir(path)

    # 確認dir
    subprocess.call('dir', shell=True)

    # インポート
    importvba()

    # pause
    os.system("pause")

if __name__ == "__main__"():
    main()
