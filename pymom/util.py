import os
import zipfile

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

                    z.extract(
                        info,
                        path=folder_name,
                        pwd=_pwd,
                    )

            except RuntimeError as e:
                print(e)
