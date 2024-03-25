import requests


class SessionManager:
    def __init__(self, url, username, password):
        self.url = url
        self.username = username
        self.password = password
        self.session = requests.Session()

    def __enter__(self):
        self.login()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.session.close()

    def get_csrf_token(self):
        response = self.session.get(self.url)
        return response.cookies["csrftoken"]

    def login(self):
        response = self.session.post(self.url, data={
            "username": self.username,
            "password": self.password,
            "csrfmiddlewaretoken": self.get_csrf_token()
        })

    def request(self, page_url):
        self.login()
        return self.session.get(page_url)
