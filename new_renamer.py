from renamer_class import Renamer
from text_art import title, error_monster, hour_glass, success_chest

#Put this data into a text file
JOB_PATH = r'J:\_Estimating Group\Jobs\ESTIMATING RESOURCE\NEW JOB FOLDER - WA'
DESTINATION_PATH = r'New Job\DWGS & SPECS\DWG\LC'
JOB_FOLDER_NAME = r'New Job'
EXCEL_FILE_NAME = r'Enterprise List.xlsx'
FILENAME_LEN_MAX = 50

TITLE = title
FATAL_ERROR = error_monster
WAITING_LOGO = hour_glass
SUCCESS_LOGO = success_chest

renamer = Renamer(JOB_PATH,
                  DESTINATION_PATH,
                  JOB_FOLDER_NAME,
                  EXCEL_FILE_NAME,
                  FILENAME_LEN_MAX,
                  TITLE,
                  FATAL_ERROR,
                  WAITING_LOGO,
                  SUCCESS_LOGO
                  )


if __name__ == '__main__':
    renamer.organize()
