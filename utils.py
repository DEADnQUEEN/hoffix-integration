from playwright.sync_api import Page, Locator
from typing import Optional

def auth(
        page: Page,
        url: str,
        selector: str,
        confirm_selector: str,
        login: str,
        password: str,
        wait_redirect_to: Optional[str] = None
) -> Page:
    page.goto(url)

    fields: Locator = page.locator(selector)

    data = [login, password]
    for index, field in enumerate(fields.all()):
        field.fill(data[index])

    page.locator(confirm_selector).first.click()

    if wait_redirect_to is None:
        return page

    page.wait_for_url(wait_redirect_to)

    return page

def get_safe_locator(page: Page, selector: str) -> Locator:
    page.wait_for_selector(selector)
    locator = page.locator(selector)
    locator.scroll_into_view_if_needed()
    return locator
