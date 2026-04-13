from sqlalchemy import text

from app.db.db import engine
from config import APP_NAME


def test_connection() -> None:
    with engine.connect() as connection:
        result = connection.execute(text("SELECT current_database();"))
        db_name = result.scalar()
        print(f"Connected to database: {db_name}")


if __name__ == "__main__":
    print(f"Starting {APP_NAME}")
    test_connection()