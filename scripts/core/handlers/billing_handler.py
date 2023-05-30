"""Importing constants"""
from scripts.constants.app_constants import DBConstants, Aggregation
from scripts.core.db.mongo.interns_b2_23.Riya.mongo_query import Item, \
    pipeline_aggregation
from scripts.utility.mongo_utility import MongoCollectionBaseClass, MongoConnect
from scripts.logging.logger import logger
from scripts.exceptions.exception_codes import *


class ItemHandler:
    """Class to handle items"""

    def __init__(self):
        self.user_mongo_obj = MongoCollectionBaseClass(database=DBConstants.DB_DATABASE,
                                                       mongo_client=MongoConnect(DBConstants.DB_URI).client,
                                                       collection=DBConstants.DB_COLLECTION)

    def read_data(self, item_id=None):
        """function to read data"""
        try:
            logger.info("Handler:read_data")
            result = self.user_mongo_obj.find({})
            if result:
                logger.info("read_data:Record Found")
                return list(result)
        except Exception as e:
            logger.error(Billing_HandlerException.EX001.format(error=str(e)))
            return {"message": "failed"}

    def create_data(self, item: Item):
        """function to create data"""
        try:
            logger.info("Handler:create_data")
            result = self.user_mongo_obj.insert_one(data=item.dict())
            if result:
                logger.info("create_data:record created successfully")
                return {"message": "success"}
        except Exception as e:
            logger.error(Billing_HandlerException.EX002.format(error=str(e)))
            return {"message": "failed"}

    def update_data(self, item_id: int, item: Item):
        """function to update data"""
        try:
            logger.info("Handler:update_data")
            result = self.user_mongo_obj.update_one({'id': item_id}, item.dict())
            if result:
                logger.info("update_data:Updated successfully")
                return {"message": "success"}
        except Exception as e:
            logger.error(Billing_HandlerException.EX003.format(error=str(e)))
            return {"message": "failed"}

    def delete_data(self, item_id: int):
        """function to delete data"""
        try:
            logger.info("Handler:delete_data")
            result = self.user_mongo_obj.delete_one({'id': item_id})
            if result:
                logger.info("delete_data:Deleted successfully")
                return {"message": "success"}
        except Exception as e:
            logger.error(Billing_HandlerException.EX004.format(error=str(e)))
            return {"message": "failed"}

    @staticmethod
    def pipeline_aggregation():
        """function for aggregation"""

        global data, aggregated_data
        try:
            logger.info("Handler:pipeline_aggregation")
            data = pipeline_aggregation(Aggregation.Agr)
            aggregated_data = []  # New list to store aggregated data
            overall_sum = 'total_amount'
            column_sum = sum(row[overall_sum] for row in data)
            aggregated_sum = {'overall_sum': column_sum}
            aggregated_data.append(aggregated_sum)
            logger.info("pipeline_aggregation: ", aggregated_data)
        except Exception as e:
            logger.error(Billing_HandlerException.EX005.format(error=str(e)))
            return {"Item Details": aggregated_data}


