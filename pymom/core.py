"""module docstring"""
import os
import zipfile
import re
from datetime import datetime, timezone

import win32com.client

INBOX = 6


class PyMom:
    def __init__(self, account: str) -> None:
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.account = self.outlook.GetNamespace("MAPI").Folders[account]

    def get_items(
        self,
        folder_path: str,
        to: str = "",
        cc: str = "",
        bcc: str = "",
        subject_contain: str = "",
        has_attachment: bool = False,
        categories: str = "",
        sent_from=None,
        sent_to=None,
    ):
        folder = self.account
        for path in folder_path.split("\\"):
            folder = folder.Folders[path]

        filtered = []
        for item in list(folder.Items):
            # 各種条件について、条件が有効かつ条件を満たさない場合はこのアイテムについての処理をスキップする
            if to and to not in item.To:
                continue
            if cc and cc not in item.CC:
                continue
            if bcc and bcc not in item.BCC:
                continue
            if subject_contain and subject_contain not in item.Subject:
                continue
            if categories and categories not in item.Categories:
                continue
            if has_attachment and not item.Attachments:
                continue
            if sent_from and sent_from > item.SentOn:
                continue
            if sent_to and sent_to < item.SentOn:
                continue

            filtered.append(item)

        # a = filtered[0]
        # keys = dir(a)
        # for k in keys:
        #     try:
        #         print(k, getattr(a, k))
        #     except:
        #         pass

        return filtered

    def move(self, condition, folder_to: str):
        # 移動先フォルダ取得
        folder = self.account
        for f in folder_to.split("\\"):
            folder = folder.Folders[f]

        # 複数アイテムを連続で処理しようとしても一度に全部処理されないことがあるので
        # 条件に合うアイテムが無くなるまでループする
        items = self.get_items(**condition)

        while len(items) != 0:
            for item in items:
                item.Move(folder)

            items = self.get_items(**condition)

    def save_message(self, condition, save_path: str):
        items = self.get_items(**condition)

        # 保存フォルダが無い場合は作成
        if not os.path.isdir(save_path):
            os.makedirs(save_path)

        # processed = []
        for item in items:
            try:
                # ファイル名に使用できない文字をエスケープ
                file_name = re.sub(r'[\\|/|:|?|.|"|<|>|\|]', "_", item.subject)
                item.SaveAs(save_path + "\\" + file_name + ".msg")
            except Exception as e:
                print(e)

    def save_attachment(
        self,
        condition,
        save_path: str,
        zip_extract: bool = False,
        zip_password: str = "",
    ):
        items = self.get_items(**condition)

        # 保存フォルダが無い場合は作成
        if not os.path.isdir(save_path):
            os.makedirs(save_path)

        for item in items:
            # 添付ファイルが無い場合は処理しない
            if not item.Attachments.Count:
                continue

            try:
                # すべての添付ファイルについて処理
                for attachment in item.Attachments:
                    file_path = save_path + "/" + attachment.FileName

                    # 同名のファイルがなければ保存
                    if not os.path.isfile(file_path):
                        attachment.SaveAsFile(file_path)

                    # zip解凍する場合の処理
                    if zip_extract:
                        extract_zip(file_path, zip_password)

            except Exception as e:
                print(e)


def extract_zip(file_path: str, password: str = ""):
    folder_name, file_name = os.path.split(file_path)
    _, ext = os.path.splitext(file_name)

    if ext == ".zip":
        # https://qiita.com/tohka383/items/b72970b295cbc4baf5ab
        with zipfile.ZipFile(file_path, "r") as z:
            try:
                for info in z.infolist():
                    # ファイル名文字化け対策
                    info.filename = info.orig_filename.encode("cp437").decode("cp932")
                    # セパレータの文字種を調整
                    if os.sep != "/" and os.sep in info.filename:
                        info.filename = info.filename.replace(os.sep, "/")

                    # zipパスワード文字列をエンコード
                    if isinstance(password, str):
                        _pwd = password.encode("utf-8")
                    else:
                        _pwd = None

                    z.extract(
                        info,
                        path=folder_name,
                        pwd=_pwd,
                    )

            except RuntimeError as e:
                print(e)


if __name__ == "__main__":
    mail_addr = ""
    myol = PyMom(mail_addr)

    subject_contain = ""
    sent_from = datetime(2022, 9, 14, tzinfo=timezone.utc)
    sent_to = datetime(2022, 9, 17, tzinfo=timezone.utc)

    condition = {
        "folder_path": "受信トレイ",
    }

    condition.update({"sent_from": sent_from})
    condition.update({"sent_to": sent_to})

    # myol.move(condition, "受信トレイ\\TEST")
    myol.save_message(condition, "P:\\")
    # myol.save_attachment(condition, "P:\\")
