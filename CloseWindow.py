# DeprecationWarning回避
# import globで呼ぶと何故かimpは推奨されてないとほざかれる
from glob import glob
import win32gui
import win32con
import time
import re

# NOTE: このファイルは元々ウィンドウが開かれている事を関係なしに作成したものなので
# 無駄な処理が書かれている可能性がある。
# その時はあなたに託します。
class CloseWindow:

    @staticmethod
    def getWindowhandle(title_name, class_name, bgr2rgb: bool = False) -> int:
        """
        引数"clsss_name", "window_name"に渡されたウィンドウのウィンドウハンドルを取得し返す。\n
        指定されたウィンドウが無かった場合はデスクトップウィンドウハンドルを返す。
        """
        # 現在アクティブなウィンドウ名を探す
        process_list = []

        def callback(handle, _):
            process_list.append(win32gui.GetWindowText(handle))
        win32gui.EnumWindows(callback, None)

        # ターゲットウィンドウ名を探す
        for process_name in process_list:
            if title_name in process_name:
                # ウィンドウハンドルを入れてブレイク
                # FindWindow(class_name(None), title_name) : https://qiita.com/SSKNOK/items/f6c09bc5eb39590f2f0b
                some_hnd = win32gui.FindWindow(class_name, process_name)
                break
            else:
                # デスクトップウィンドウのハンドルを取得
                some_hnd = win32gui.GetDesktopWindow()

        # NOTE: このクラス以外にも使う可能性があるので一旦返す
        return some_hnd

    def closeWindow(self, someHandle:int):
        """
        渡されたウィンドウハンドルのウィンドウを閉じる。
        """
        # 指定したウィンドウを最前面に表示させる
        # winHndに入っているハンドルがデスクトップでなければ実行
        if self.desktopWinHnd != someHandle:
            win32gui.SetWindowPos(someHandle,win32con.HWND_TOPMOST,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        # ウィンドウハンドルを使用して閉じるボタンをクリックする
        win32gui.PostMessage(someHandle,win32con.WM_CLOSE,0,0)

        # ウィンドウが閉じたら次の処理が進むようにする。
        # デスクトップウィンドウハンドルが帰って来ない限りFalseになる
        #  ↳指定したウィンドウが閉じない限り永遠にループし続ける
        winhandleType = self.checkwindow(someHandle)
        while winhandleType:
            time.sleep(0.5)
            someHandle = self.getWindowhandle(self.title_name, self.class_name)
            winhandleType = self.checkwindow(someHandle)

    def checkwindow(self, winHnd: int):
        """
        渡されたウィンドウハンドルがデスクトップウィンドウハンドルであればTrue\n
        そうでなければFalseを返す。
        """
        if winHnd != self.desktopWinHnd:
            winhandleType = True
        else:
            winhandleType = False

        return winhandleType

# デザインパターン、strategy
# 一連のアルゴリズムを実装する
class Strategy(CloseWindow):
    """
    CloseWindowクラスを継承して指定したウィンドウを閉じる一連の動作を行う。\n
    インスタンスを関数として扱うと処理が始まる。
    """

    def __init__(self) -> None:
        # binフォルダに入ってるファイルを解析してtitle_name, class_nameを読み込む
        self.title_name, ext = self.getXlsmName()
        # クラスネームを取得
        self.class_name = self.getClassName(ext)
        # デスクトップウィンドウのハンドルを取得
        self.desktopWinHnd = win32gui.GetDesktopWindow()

    # __call__:インスタンス名を関数のように呼び出したら処理される内容
    def __call__(self):

        return self.close()

    def getXlsmName(self) -> str:
        """
        binフォルダに入っているファイルの名前と拡張子を取得する。
        ext(extension): 拡張子
        return filename, ext
        """
        # DeprecationWarning回避
        # globモジュールのメソッドである
        path = glob("bin/*")

        # 正規表現
        # \w : 英数字のみ取得する
        p = r"\w+"
        # re.findall(p, s): 重複しないマッチを文字列のリストに入れる
        bin, filename, ext = re.findall(p, path[0])
        # binは使わない
        del bin

        # テスト
        #print(f"ファイル名:{filename}\n拡張子:.{ext}")

        return filename, ext

    # NOTE: 将来的に色んなクラスネームを検索するかもしれないので一応用意しておく
    def getClassName(self, ext) -> str:
        """
        拡張子から登録されているクラスネームを取得する
        """
        ext_list = {"xlsm":"XLMAIN"}
        class_name = ext_list[ext]

        return class_name

    def close(self):
        """
        指定したウィンドウを閉じる。\n
        閉じたらTrueが返ってくる。
        """
        someHandle = self.getWindowhandle(self.title_name, self.class_name)

        print(f"ウィンドウハンドル:{someHandle}\nデスクトップ:{not self.checkwindow(someHandle)}")

        print('\033[31m' + "Error : ファイルが開かれているためインポート出来ません。\n保存してウィンドウを閉じて下さい。" + '\033[0m')

        # 基底のメソッド
        self.closeWindow(someHandle)

        return True


if __name__ == "__main__":
    obj = Strategy()
    obj()