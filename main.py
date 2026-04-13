from sqlalchemy import text

from app.db.db import engine
from config import APP_NAME


def test_connection() -> None:
    with engine.connect() as connection:
        current_db = connection.execute(text("SELECT current_database();")).scalar()
        current_user = connection.execute(text("SELECT current_user;")).scalar()

        print(f"Application: {APP_NAME}")
        print(f"Connected to database: {current_db}")
        print(f"Connected as user: {current_user}")


if __name__ == "__main__":
    test_connection()