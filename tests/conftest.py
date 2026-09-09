import pytest

from report_pipeline.config import Settings


@pytest.fixture
def settings() -> Settings:
    return Settings(
        smtp_host="smtp.example.com",
        smtp_port=587,
        smtp_user="sender@example.com",
        smtp_password="secret",
        mail_from="sender@example.com",
        mail_to="me@company.example",
        report_tz="America/New_York",
    )
