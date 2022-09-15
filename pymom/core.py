"""module docstring"""
import os
import win32com.client

from mail import OutlookMail
import time

INBOX = 6


class PyMom:
    def __init__(self, account: str) -> None:
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.account = self.outlook.GetNamespace("MAPI").Folders[account]
        # self.inbox = self.outlook.GetDefaultFolder(INBOX)

    def get_items(self, folder_path: str, condition: dict = None):
        folders = self.account
        for f in folder_path.split("\\"):
            folders = folders.Folders[f]
        return folders.Items

    def move(self, items, folder_to: str):
        # 移動先フォルダ取得
        folder = self.account
        for f in folder_to.split("\\"):
            folder = folder.Folders[f]

        for item in items:
            item.Move(folder)
            time.sleep(1)

        # TODO: Gmailと同期しているOutlookの場合、同期処理が走ってフォルダ内の全メールを一度に移動できない


if __name__ == "__main__":
    mail_addr = ""
    myol = PyMom(mail_addr)

    path = "受信トレイ"
    mails = myol.get_items(path)

    print(len(mails))

    for m in mails:
        print(m.subject)

    myol.move(mails, "受信トレイ\\TEST")
