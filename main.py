"""手寫估價單自動記帳工具。

雙擊執行本檔案，或在命令列執行：python main.py
所有設定都可在程式畫面中調整；不需要安裝額外 Python 套件。
"""

from bookkeeping import BookkeepingApp


if __name__ == "__main__":
    BookkeepingApp().run()
