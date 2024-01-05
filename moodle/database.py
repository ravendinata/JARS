import pymysql
from termcolor import colored

import config

@staticmethod
def test_connection(supress = True):
    """
    Tests the connection to the moodle database.
    
    Returns:
        bool: Whether the connection was successful.

    Raises:
        ValueError: If the moodle database credentials are not set.
    """
    # Check if moodle database credentials are set
    config_keys = ["moodle_db_server", "moodle_db_username", "moodle_db_name", "moodle_db_password"]
    if any(config.get_config(key) is None or config.get_config(key) == "" for key in config_keys):
        raise ValueError("Moodle database credentials and/or server information not set! Please check configuration or contact administrator.")

    try:
        with pymysql.connect(
            host = config.get_config("moodle_db_server"),
            user = config.get_config("moodle_db_username"),
            password = config.get_config("moodle_db_password"),
            db = config.get_config("moodle_db_name"),
            charset = "utf8mb4",
            cursorclass = pymysql.cursors.DictCursor
        ) as conn:
            with conn.cursor() as cursor:
                sql = "SHOW TABLES;"
                cursor.execute(sql)
                result = cursor.fetchall()
                if not supress:
                    print(result)
    except Exception as e:
        print(colored(f"Exception: {e}", "red"))
        print(colored("Connection to Moodle database failed!", "red"))
        return False

    return True

class Database:
    def __init__(self):
        self.conn = pymysql.connect(
            host = config.get_config("moodle_db_server"),
            user = config.get_config("moodle_db_username"),
            password = config.get_config("moodle_db_password"),
            db = config.get_config("moodle_db_name"),
            charset = "utf8mb4",
            cursorclass = pymysql.cursors.DictCursor
        )

    def __del__(self):
        self.conn.close()

    def query(self, sql):
        with self.conn.cursor() as cursor:
            cursor.execute(sql)
            result = cursor.fetchall()
            return result
        
    def query_from_file(self, filename, verbose = False, **kwargs):
        with open(filename, "r") as f:
            sql = f.read()
            sql = sql.format(**kwargs).replace("\n", " ").replace("  ", " ")
            if verbose:
                print(sql)
            return self.query(sql)
    
    # Simple data queries (no joins)
    def get_users(self, columns = "username", filter = None):
        if filter is None:
            return self.query(f"SELECT {columns} FROM mdl_user;")
        else:
            return self.query(f"SELECT {columns} FROM mdl_user WHERE {filter};")
        
    def get_courses(self, columns = "id, fullname", filter = None):
        if filter is None:
            return self.query(f"SELECT {columns} FROM mdl_course;")
        else:
            return self.query(f"SELECT {columns} FROM mdl_course WHERE {filter};")
        
    def get_cohorts(self, columns = "id, name", filter = None):
        if filter is None:
            return self.query(f"SELECT {columns} FROM mdl_cohort;")
        else:
            return self.query(f"SELECT {columns} FROM mdl_cohort WHERE {filter};")
