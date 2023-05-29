"""Importing MongoClient for connection"""
from fastapi import FastAPI
from pymongo import MongoClient
from scripts.constants.app_constants import DBConstants
from scripts.core.schema.model import Item
from scripts.exceptions.exception_codes import Mongo_queryException
from scripts.logging.logger import logger
import xlsxwriter

app = FastAPI()
client = MongoClient(DBConstants.DB_URI)
db = client[DBConstants.DB_DATABASE]
# # Creating document
billing = db[DBConstants.DB_COLLECTION]

# excel = billing.find({})
# sheet = list(excel)
#
#
#
# # Create a new Excel file
# workbook = xlsxwriter.Workbook('reports/billing_report.xlsx')
# worksheet = workbook.add_worksheet()
#
# # Define cell formats
# bold_format = workbook.add_format({'bold': True})
# red_format = workbook.add_format({'bg_color': 'red'})
#
# # Write headers
# headers = list(sheet[0].keys())
# for col, header in enumerate(headers):
#     worksheet.write(0, col, header, bold_format)
#
# # Write bill data to the worksheet
# for row, data in enumerate(sheet, start=1):
#     for col, value in enumerate(data):
#         if col == len(data) - 1:
#             try:
#                 if int(value) > 1000:  # Convert value to integer
#                     worksheet.write(row, col, int(value), red_format)  # Convert value to integer
#                 else:
#                     worksheet.write(row, col, value)
#             except ValueError:
#                 worksheet.write(row, col, value)  # Handle non-numeric values
#         else:
#             worksheet.write(row, col, value)
#
# # Auto-fit column widths
# worksheet.autofilter(0, 0, row, col)
# for col in range(len(headers)):
#     worksheet.set_column(col, col, 15)
#
# # Close the workbook
# workbook.close()
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
