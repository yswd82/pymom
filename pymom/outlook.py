import win32com.client
import os
import zipfile


class OutlookWatcher:
    def __init__(self, account):
        self.outlook = win32com.client.Dispatch("Outlook.Application")

        self._account = self.outlook.GetNamespace("MAPI").Folders[account]

    def get_items(self, path=""):
        """pathで指定した階層のメールを取得する
        Args:
            path (str, optional): メールフォルダのパスを/区切りで指定する Defaults to "".
        """

        if "/" in path:
            path = path.split("/")
        else:
            path = [path]

        # フォルダを初期化
        folder = self._account

        # pathで指定したフォルダを取得する
        for p in path:
            folder = folder.Folders[p]

        return folder.Items

    def send_mail(self, to="", cc="", bcc="", subject="", bodyformat=1, body=""):
        mail = self.outlook.CreateItem(0)
        mail.to = to
        mail.cc = cc
        mail.bcc = bcc
        mail.subject = subject
        mail.bodyFormat = bodyformat
        mail.body = body

        mail.display(True)


class Mail:
    """メールアイテム"""

    def __init__(self, outlookmail):
        self.mail = outlookmail

    def save_attachment(
        self, save_path: str, zip_extract: bool = False, zip_password: str = None
    ):
        """添付ファイルを保存する
        Args:
            save_path (str): 保存パス
            zip_extract (bool, optional): zipを展開するか. Defaults to False.
            zip_password (str, optional): zipのパスワード. Defaults to None.
        Returns:
            list: 保存したファイルのリスト
        """

        # 添付ファイルが無い場合はリターン
        if not self.mail.Attachments.Count:
            return

        # 保存フォルダが無い場合は作成
        if not os.path.isdir(save_path):
            os.makedirs(save_path)

        res = []
        try:
            # すべての添付ファイルについて処理
            for attachment in self.mail.Attachments:
                file_path = save_path + "/" + attachment.FileName

                # 同名のファイルがなければ保存
                if not os.path.isfile(file_path):
                    attachment.SaveAsFile(file_path)
                    # 保存できた場合はフルパスを返す
                    res.append(file_path)

                # zip解凍する場合の処理
                if zip_extract:
                    self.extract_zip(file_path, zip_password)

        except Exception as e:
            print(e)

        return res

    @staticmethod
    def extract_zip(file_path: str, password: str = None):
        """zipを展開する
        Args:
            file_path (str): フルパスのファイル名
            password (str, optional): zipパスワード. Defaults to None.
        """
        foldername, filename = os.path.split(file_path)
        filename_body, ext = os.path.splitext(filename)

        if ext == ".zip":
            # https://qiita.com/tohka383/items/b72970b295cbc4baf5ab
            with zipfile.ZipFile(file_path, "r") as z:
                try:
                    for info in z.infolist():
                        # ファイル名文字化け対策
                        info.filename = info.orig_filename.encode("cp437").decode(
                            "cp932"
                        )
                        # セパレータの文字種を調整
                        if os.sep != "/" and os.sep in info.filename:
                            info.filename = info.filename.replace(os.sep, "/")

                        # zipパスワード文字列をエンコード
                        if isinstance(password, str):
                            _pwd = password.encode("utf-8")
                        else:
                            _pwd = None

                        # 展開
                        z.extract(
                            info,
                            path=foldername,
                            pwd=_pwd,
                        )

                except RuntimeError as e:
                    print(e)

if __name__ == "__main__":
    myou = OutlookWatcher("Sawada_Yousuke@smtb.jp")
    mails = myou.get_items("受信トレイ")

    _max = 0
    for _ in mails:
        # 受信トレイ内のメールを一覧
        print(_.sendername, _.senderemailaddress, _.subject)

        if _max > 100:
            exit()
        _max+=1

        # # 件名が"attachtest2"のメールの添付ファイル保存&zip展開(パス付き)
        # if _.subject == "attachtest2":
        #     mymail = Mail(_)
        #     res = mymail.save_attachment(
        #         save_path="E:\\save", zip_extract=True, zip_password="1234"
        #     )
        #     print(_.subject, "saved", res)

        # # 件名が"attachtest3"のメールの添付ファイル保存&zip展開
        # if _.subject == "attachtest3":
        #     mymail = Mail(_)
        #     res = mymail.save_attachment(save_path="E:\\save", zip_extract=True)
        #     print(_.subject, "saved", res)