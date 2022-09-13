import win32com.client


class PyMom:
    INBOX = 6

    def __init__(self):
        self.outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
        self.inbox = self.outlook.GetDefaultFolder(self.INBOX)

    def move_mail_item(self, condition:dict, to_folder:list, from_folder:list=None):
        # 先フォルダ取得
        to_folders =self.inbox.folders
        for f in to_folder:
            to_folders = to_folders(f)

        # 元フォルダ取得
        if from_folder:
            from_folders =  self.inbox.folders
            for f in from_folder:
                from_folders = from_folders(f)
            messages = from_folders.Items
        else:
            # 指定ない場合は受信トレイ
            messages = self.inbox.Items

        # 仕分けを行う
        for message in messages:
            if   'パイナップル' in message.subject:
                message.Move(to_folders)
                print(message)
                print(dir(message))
                keys = dir(message)

                for k in keys:
                    print(k, getattr(message,k))


if __name__ == '__main__':
    myol = PyMom()

    myol.move_mail_item({},['901k'])