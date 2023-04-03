import os
import pandas as pd
from pymongo import MongoClient

DB_NAME = ""
URI = "mongodb://localhost:27017"
DIR_PATH = r""


class MongoDb:
    def __init__(self):
        try:
            self.client = MongoClient(URI)
            self.database = self.client[DB_NAME]
            self.__isInitialized = True
        except Exception as e:
            self.__isInitialized = False
            self.close()
            print(e)

    def is_connected(self) -> bool:
        return self.__isInitialized

    def insert(self, collection: str, data: dict[str, any]) -> str:
        return self.database[collection].insert_one(data).inserted_id

    def insert_many(self, collection: str, data: list[dict[str, any]]):
        return self.database[collection].insert_many(data).inserted_ids

    def find(self, collection, query):
        return self.database[collection].find(query)

    def find_all(self, collection):
        return self.database[collection].find({})

    def find_one(self, collection, query):
        return self.database[collection].find_one(query)

    def delete(self, collection, query):
        return self.database[collection].delete_one(query)

    def update(self, collection, query, data):
        return self.database[collection].update_one(query, data)

    def close(self) -> bool:
        try:
            self.client.close()
            return True
        except Exception as e:
            print(e)
            return False


def get_paths_and_name(dir: str) -> tuple[list[str], list[str]]:
    paths = []
    names = []
    for path in os.listdir(dir):
        if path.endswith(".xlsx"):
            paths.append(os.path.join(dir, path))
            # Append the name of the file without the extension
            name = path.split("\\")
            names.append(name[len(name) - 1].split(".")[0])
    return paths, names


def get_schema(df: pd.DataFrame) -> dict[str, any]:
    schema = {}
    for col in df.columns:
        schema[col] = None
    return schema


def index_collection_names(names: list[str]):
    mongo = MongoDb()
    if mongo.is_connected():
        mongo.insert("general", {"collection_names": names})
        mongo.close()


def main(paths: list[str], names: list[str]):
    index = 0
    for path in paths:
        df = pd.read_excel(path)
        list_of_schemas: list[dict[str, any]] = []
        for _, row in df.iterrows():
            list_of_schemas.append(row.to_dict())

        mongo = MongoDb()
        if mongo.is_connected():
            mongo.insert_many(names[index], list_of_schemas)
            mongo.close()
        index += 1

    index_collection_names(names)


if __name__ == "__main__":
    paths, names = get_paths_and_name(DIR_PATH)
    main(paths, names)
