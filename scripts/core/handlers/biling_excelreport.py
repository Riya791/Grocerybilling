import xlsxwriter
from fastapi import FastAPI
from scripts.constants.app_constants import DBConstants, Aggregation
from scripts.core.db.mongo.interns_b2_23.Riya.mongo_query import billing
from scripts.core.handlers.billing_handler import pipeline_aggregation
from scripts.utility.mongo_utility import MongoCollectionBaseClass, MongoConnect
from scripts.exceptions.exception_codes import Excel_billingException
from scripts.logging.logger import logger

app = FastAPI()


class ExcelGeneration:
    """class for generating excel"""

    def __init__(self):
        self.mongo_obj = MongoCollectionBaseClass(database=DBConstants.DB_DATABASE,
                                                  mongo_client=MongoConnect(DBConstants.DB_URI).client,
                                                  collection=DBConstants.DB_COLLECTION)

    @staticmethod
    def excel_billing():
        global row, col
        try:
            logger.info("Handler:excel_billing")
            excel_data = pipeline_aggregation(Aggregation.Agr)
            print(excel_data)

            excel = billing.find({})
            sheet = list(excel)

            # Create a new Excel file
            workbook = xlsxwriter.Workbook('Report/billing_report.xlsx')
            worksheet = workbook.add_worksheet()

            # Define cell formats
            bold_format = workbook.add_format({'bold': True})
            red_format = workbook.add_format({'bg_color': 'red'})

            # Write headers
            headers = list(sheet[0].keys())
            for col, header in enumerate(headers):
                worksheet.write(0, col, header, bold_format)

            # Write bill data to the worksheet
            for row, data in enumerate(sheet, start=1):
                for col, value in enumerate(data):
                    if col == len(data) - 1:
                        try:
                            if int(value) > 1000:  # Convert value to integer
                                worksheet.write(row, col, int(value), red_format)  # Convert value to integer
                            else:
                                worksheet.write(row, col, value)
                        except ValueError:
                            worksheet.write(row, col, value)  # Handle non-numeric values
                    else:
                        worksheet.write(row, col, value)

            # Auto-fit column widths
            worksheet.autofilter(0, 0, row, col)
            for col in range(len(headers)):
                worksheet.set_column(col, col, 15)

            # Close the workbook
            workbook.close()
        except Exception as e:
            logger.error(Excel_billingException.EX0020.format(error=str(e)))
