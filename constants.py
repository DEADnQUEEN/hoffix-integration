LOGIN_JSON = {
    "grant_type": "password",
    "client_id": "hoff-web-app-prod",
    "client_secret": "Th1s0Is2supAS3cr9t",
    "app_version": "1.16.0.0",
}

# set: get
# '.' - devider
MAPPING_JSON = {
    "number": "number",
    "serviceId": "service.id",
    "registryTypeId": "registryType.id",
    "serviceDetail": "serviceDetail",
    "workDate": "workDate",
    "workTime": "workTime",
    "clientName": "clientName",
    "clientPhone": "clientPhone",
    "address": "address",
    "comment": "comment",
    "contractorId": "contractor.id",
    "workerId": "worker.id",
    "orderStatus": "orderStatus",
    "workerAssistantIds": [],
}

FROM_DB_COLLECT = {
    "workerId": ("https://back-hoffix.hoff.ru/api/Contractors", "workers", False, lambda data: [{"name": row['fullName'].strip(), "id":  row['id']} for row in data[0]['workers']]),
    "registryTypeId": ("https://back-hoffix.hoff.ru/api/RegistryTypes?pageNumber=1&pageSize=10000&includeArchived=true", "registry", False)
}

SKIP_FIELDS = {
    "workerId": lambda field: field.lower().strip() == "отмена"
}

OUTPUT_LIST = "Протокол"
DATE_FORMAT = "%d-%m-%Y %H:%M:%S"

