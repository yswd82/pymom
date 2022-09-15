import os
import zipfile


class OutlookMail:
    def __init__(self, mail):
        self._mail = mail

    def save_attachment(
        self, save_path: str, extract_zip: bool = False, zip_password: str = ""
    ):
        # 添付ファイルが無い場合はリターン
        if not self._mail.Attachments.Count:
            return

        # 保存フォルダが無い場合は作成
        if not os.path.isdir(save_path):
            os.makedirs(save_path)

        res = []
        try:
            # すべての添付ファイルについて処理
            for attachment in self._mail.Attachments:
                file_path = save_path + "/" + attachment.FileName

                # 同名のファイルがなければ保存
                if not os.path.isfile(file_path):
                    attachment.SaveAsFile(file_path)
                    # 保存できた場合はフルパスを返す
                    res.append(file_path)

                # zip解凍する場合の処理
                if extract_zip:
                    self.extract_zip(file_path, zip_password)

        except Exception as e:
            print(e)

        return res

    @staticmethod
    def extract_zip(file_path: str, password: str = ""):
        folder_name, file_name = os.path.split(file_path)
        _, ext = os.path.splitext(file_name)

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
                            path=folder_name,
                            pwd=_pwd,
                        )

                except RuntimeError as e:
                    print(e)
