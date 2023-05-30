"""Classes for handling constants"""


class APis:
    # create_api = '/create'
    view_all_items_api = '/view_all_data'
    create_api = '/items/'
    update_api = '/items/{items_id}'
    delete_api = '/delete/{items_id}'
    send_email = '/send_email'
    get_api = '/billing-price'
    get_excel = '/billing_report'


class DBConstants:
    DB_URI = 'mongodb://intern_23:intern%40123@192.168.0.220:2717/interns_b2_23'
    DB_DATABASE = 'interns_b2_23'
    DB_COLLECTION = 'riya_billing'


class Aggregation:
    Agr = [
        {
            '$addFields': {
                'total_amount': {
                    '$multiply': [
                        '$quantity', '$cost'
                    ]
                }
            }
        }, {
            '$project': {
                '_id': 0
            }
        }
    ]


class CommonConstants:
    APP_KEY = "main:app"
