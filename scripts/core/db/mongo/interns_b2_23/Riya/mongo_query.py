"""Importing MongoClient for connection"""
from fastapi import FastAPI
from pymongo import MongoClient
from scripts.constants.app_constants import DBConstants
from scripts.core.schema.model import Item
from scripts.exceptions.exception_codes import Mongo_queryException
from scripts.logging.logger import logger

app = FastAPI()
client = MongoClient(DBConstants.DB_URI)
db = client[DBConstants.DB_DATABASE]
# # Creating document
billing = db[DBConstants.DB_COLLECTION]


# creating class


def read_item():
    """Function to read items"""
    logger.info("mongo_query:read_item")
    data = []
    try:
        for items in billing.find():
            del items['_id']
            data.append(items)
    except Exception as e:
        logger.error(Mongo_queryException.EX0015.format(error=str(e)))
    return {
            "db": data
        }


def create_item(item: Item):
    """Function to create item"""
    try:
        logger.info("mongo_query:create_item")
        billing.insert_one(item.dict())
        db[item.id] = item.name
    except Exception as e:
        logger.error(Mongo_queryException.EX0016.format(error=str(e)))
    return {
        "db": db
    }


def update_item(item_id: int, item: Item):
    """Function to update item"""
    try:
        logger.info("mongo_query:update_item")
        billing.update_one({"id": item_id}, {"$set": item.dict()})
    except Exception as e:
        logger.error(Mongo_queryException.EX0017.format(error=str(e)))


def delete_item(item_id: int):
    """Function to delete item"""
    try:
        logger.info("mongo_query:delete_item")
        billing.delete_one({"id": item_id})
    except Exception as e:
        logger.error(Mongo_queryException.EX0018.format(error=str(e)))
    return {"message": "deleted"}


def pipeline_aggregation(pipeline: list):
    """Function for aggregation"""
    try:
        logger.info("mongo_query:pipeline_aggregation")
    except Exception as e:
        logger.error(Mongo_queryException.EX0019.format(error=str(e)))
    return billing.aggregate(pipeline)
