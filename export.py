import os
import subprocess

def main():
    #自分の相対パス取得
    path = os.path.dirname(__file__)

    #自分のファイルパスにcd
    os.chdir(path)

    #確認dir
    subprocess.call('dir', shell=True)

    #エクスポート
    print("\nエクスポートを実行します。\n")
    subprocess.call('cscript vbac.wsf decombine', shell=True)

    #pause
    os.system("pause")

if __name__ == "__main__"():
    main()
