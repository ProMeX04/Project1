from session_manager import *
from dataframe_manager import *
from excel_manager import *
from config_manager import *

if __name__ == "__main__":
    with SessionManager(URL, USERNAME, PASSWORD) as session:
        response = session.request(PAGE_URL)
    if response.status_code == 200:
        print("Connected")
        dataframe = DataFrameManager(
            response, CURRENT_LESSON, LESSON, KEYWORD).new_df()
        excel = ExcelManager()
        excel.conv2excel(dataframe, CURRENT_LESSON, LESSON)
        print("done!")
        excel.save_and_open(FILE_PATH, CLASS)
