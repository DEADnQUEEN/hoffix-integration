import requests
import json
import constants
from typing import Optional, Callable


class HoffixAPI:
    @staticmethod
    def __create_login_json(login, password):
        return {
            **constants.LOGIN_JSON,
            "username": login,
            "password": password,
        }

    @staticmethod
    def __get_auth_token(login, password):
        auth_request = requests.post(
            "https://id-hoffix.hoff.ru/connect/token",
            data=HoffixAPI.__create_login_json(login, password),
        )

        auth = json.loads(auth_request.text)
        return f"{auth['token_type']} {auth['access_token']}"

    def __init__(self, login, password):
        self.token = self.__get_auth_token(login, password)
        self.headers = {
            "Authorization": self.token,
            "Content-Type": "application/json"
        }

        self.__cached_values = {}

    def find_order(self, search_by, date: str, default: any = None):
        order_dataset = requests.get(
            f"https://hoffix-external-bff.prod-omni.hoff.ru/v1/visit/order?search={search_by}&work-date-from={date}&work-date-to={date}",
            headers=self.headers,
        )

        orders = json.loads(order_dataset.text)["data"].get("visitOrders", [])

        if len(orders) == 0:
            return default
        return orders[0]

    def get_full_order_data(self, order_id):
        order_data = requests.get(
            f"https://back-hoffix.hoff.ru/api/Orders/{order_id}",
            headers=self.headers,
        )
        return json.loads(order_data.text)['data']

    def set_order_data(self, order_id, order_data):
        request = requests.put(
            f"https://back-hoffix.hoff.ru/api/Orders/{order_id}",
            headers=self.headers,
            json=order_data,
        )

        return json.loads(request.text)

    def __get_from_hoffix_db(self, url: str, cached_field: Optional[str] = None, clear_cache: bool = False):
        if cached_field is not None and not clear_cache:
            cache = self.__cached_values.get(cached_field)
            if cache is not None:
                return cache

        request = requests.get(url, headers={"Authorization": self.token})
        data = json.loads(request.text)['data']

        if cached_field is not None:
            self.__cached_values[cached_field] = data

        return data

    def get_dataset_from_db(self, url: str, cached_field: Optional[str] = None, clear_cache: bool = False, edit_function: Optional[Callable[[list[dict[str, any]]], list[dict[str, any]]]] = None) -> dict[str, any]:
        dataset = self.__get_from_hoffix_db(url, cached_field, clear_cache)

        if edit_function is not None:
            dataset = edit_function(dataset)

        return {
            row['name']: row['id']
            for row in dataset
        }
