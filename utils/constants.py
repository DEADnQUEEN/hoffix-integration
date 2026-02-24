LOGIN_JSON = {
    "grant_type": "password",
    "client_id": "hoff-web-app-prod",
    "client_secret": "Th1s0Is2supAS3cr9t",
    "app_version": "1.16.0.0",
}

SKIP_FIELDS = {
    "workerId": lambda field: field.lower().strip() == "отмена"
}
