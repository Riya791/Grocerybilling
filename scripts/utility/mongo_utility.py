"""Importing MongoClient"""
from typing import Dict, Optional
from pymongo import MongoClient
from pymongo.cursor import Cursor

"""Classes for Mongo connections"""


class MongoConnect:
    def __init__(self, uri):
        try:
            self.uri = uri
            self.client = MongoClient(self.uri, connect=False)
        except Exception as e:
            print(f"Exception in connection {(str(e))}")
            raise e

    def __call__(self, *args, **kwargs):
        return self.client


class MongoCollectionBaseClass:
    def __init__(self, mongo_client, database, collection):
        self.client = mongo_client
        self.database = database
        self.collection = collection
        # print(self.database, self.collection)

    def insert_one(self, data: Dict):
        try:
            database_name = self.database
            collection_name = self.collection
            db = self.client[database_name]
            collection = db[collection_name]
            response = collection.insert_one(data)
            return response.inserted_id
        except Exception as e:
            print(f"Error in inserting the data {str(e)}")
            raise e

    def find(
            self,
            query: Dict,
            filter_dict: Optional[Dict] = None,
            sort=None,
            skip: Optional[int] = 0,
            collation: Optional[bool] = False,
            limit: Optional[int] = None,
    ) -> Cursor:
        """
        The function is used to query documents from a given collection in a Mongo Database
        :param query: Query Dictionary
        :param filter_dict: Filter Dictionary
        :param sort: List of tuple with key and direction. [(key, -1), ...]
        :param skip: Skip Number
        :param limit: Limit Number
        :param collation:
        :return: List of Documents
        """
        if sort is None:
            sort = list()
        if filter_dict is None:
            filter_dict = {"_id": 0}
        database_name = self.database
        collection_name = self.collection
        try:
            db = self.client[database_name]
            collection = db[collection_name]
            if len(sort) > 0:
                cursor = (
                    collection.find(
                        query,
                        filter_dict,
                    )
                    .sort(sort)
                    .skip(skip)
                )
            else:
                cursor = collection.find(
                    query,
                    filter_dict,
                ).skip(skip)
            if limit:
                cursor = cursor.limit(limit)
            if collation:
                cursor = cursor.collation({"locale": "en"})

            return cursor
        except Exception as e:
            print(f"Error in fetching {str(e)}")
            raise e

    def update_one(
            self,
            query: Dict,
            data: Dict,
            upsert: bool = False,
    ):
        """

        :param upsert:
        :param query:
        :param data:
        :return:
        """
        try:
            database_name = self.database
            collection_name = self.collection
            db = self.client[database_name]
            collection = db[collection_name]
            response = collection.update_one(query, {"$set": data}, upsert=upsert)
            return response.modified_count
        except Exception as e:
            print(f"Failed to update one doc {str(e)}")
            raise e

    def delete_one(self, query: Dict):
        """
        :param query:
        :return:
        """
        try:
            database_name = self.database
            collection_name = self.collection
            db = self.client[database_name]
            collection = db[collection_name]
            response = collection.delete_one(query)
            return response.deleted_count
        except Exception as e:
            print(f"Failed to delete {str(e)}")
            raise e

    def count_documents(self, query: Dict):
        """
        :param query:
        :return:
        """
        try:
            database_name = self.database
            collection_name = self.collection
            db = self.client[database_name]
            collection = db[collection_name]
            response = collection.count_documents(query)
            return response
        except Exception as e:
            print(f"Failed to count documents {str(e)}")
            raise e
