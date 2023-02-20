from OrdersManager.storage import HandleStorage


class StorageProxy:

    def __init__(self, storage_file):
        self.storage = HandleStorage(storage_file)
