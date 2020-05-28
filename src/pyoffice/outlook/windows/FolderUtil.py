__all__ = []


class FolderUtil:

    @staticmethod
    def hasFolderExists(folder,
                        folderName: str):
        return folderName in folder.getFolderNameList()
