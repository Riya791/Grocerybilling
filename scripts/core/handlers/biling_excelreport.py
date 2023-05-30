import xlsxwriter
from fastapi import FastAPI
from fastapi.responses import FileResponse
from scripts.constants.app_constants import DBConstants, Aggregation
from scripts.core.db.mongo.interns_b2_23.Riya.mongo_query import billing
from scripts.core.handlers.billing_handler import pipeline_aggregation
from scripts.exceptions.exception_codes import Excel_billingException
from scripts.logging.logger import logger
from scripts.utility.mongo_utility import MongoCollectionBaseClass, MongoConnect

app = FastAPI()


class ExcelGeneration:
    """class for generating excel"""

    def __init__(self):
        self.mongo_obj = MongoCollectionBaseClass(database=DBConstants.DB_DATABASE,
                                                  mongo_client=MongoConnect(DBConstants.DB_URI).client,
                                                  collection=DBConstants.DB_COLLECTION)

    @staticmethod
    def excel_billing():
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
            worksheet.set_row(0, None, bold_format)

            # Auto-adjust column width
            worksheet.set_column(0, 3, 15)

            # Write column headers
            headers = ["id", "name", "quantity", "cost"]
            for col, header in enumerate(headers):
                worksheet.write(0, col, header)

            # Write data to the worksheet
            for row, data in enumerate(sheet):
                worksheet.write(row + 1, 0, data.get("id"))
                worksheet.write(row + 1, 1, data.get("name"))
                worksheet.write(row + 1, 2, data.get("quantity"))
                worksheet.write(row + 1, 3, data.get("cost"))

            workbook.close()
            return FileResponse("Report/billing_report.xlsx", filename="billing_report.xlsx")
        except Exception as e:
            logger.error(Excel_billingException.EX0020.format(error=str(e)))
