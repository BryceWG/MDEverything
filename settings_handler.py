from PyQt6.QtCore import QSettings

class SettingsHandler:
    def __init__(self, company, app_name):
        self.settings = QSettings(company, app_name)

    def save_setting(self, key, value):
        self.settings.setValue(key, value)

    def load_setting(self, key, default=None):
        return self.settings.value(key, default)

    def save_pdf_settings(self, app_id, secret_code):
        self.save_setting("app_id", app_id)
        self.save_setting("secret_code", secret_code)

    def load_pdf_settings(self):
        return {
            "app_id": self.load_setting("app_id", ""),
            "secret_code": self.load_setting("secret_code", "")
        }