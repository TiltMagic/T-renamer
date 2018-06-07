import os
import shutil
import xlsxwriter


class Renamer:

    # path_pointer = None
    # user_input = None
    # root_path = None

    def __init__(self,
                 job_path,
                 destination_path,
                 job_folder_name,
                 excel_file_name,
                 filename_len_max,
                 title,
                 fatal_error,
                 waiting_logo,
                 success_logo):
                 
        self.job_path = job_path
        self.destination_path = destination_path
        self.job_folder_name = job_folder_name
        self.excel_file_name = excel_file_name
        self.filename_len_max = filename_len_max
        self.title = title
        self.fatal_error = fatal_error
        self.waiting_logo = waiting_logo
        self.success_logo = success_logo

        try:
            self.root_path = os.getcwd()
            # self.path_pointer = self.root_path
        except:
            print('Unable to build root path and pointer')
            print(self.fatal_error)

    def display_greeting(self):
        print("Welcome to Tilley Renamer v.3.2\n")
        print(self.title)


    def clear_terminal(self):
        os.system('cls' if os.name == 'nt' else 'clear')
        self.display_greeting()

    def get_user_input(self):
        while True:
            self.clear_terminal()
            user_input = input('''
    OPTIONS:\n
    A) For Numbered Files:
       Enter a starting number for you files, then press 'ENTER'\n

    B) For NON-numbered Files:
       Just press 'ENTER'\n
    >> ''').strip()
            if user_input != '':
                try:
                    int(user_input)
                    print(self.waiting_logo)
                    print('''...wait a couple of seconds for the magic to happen''')
                    self.user_input = user_input
                    return user_input
                except ValueError:
                    self.clear_terminal()
                    print("Please enter a valid number...\n")
                    input("Press 'ENTER' to continue")
            else:
                self.user_input = user_input
                return user_input


    def rename_files_numbered(self, file_path, filename_len_max, user_input):
        '''Renames files and adds count to title'''
        num = int(user_input)
        file_format = '{}-{}'
        for _file in os.listdir(file_path):
            path, ext = os.path.splitext(_file)
            if ext == '.pdf':
                if len(_file) >= filename_len_max:
                    os.rename(_file, file_format.format(str(num).zfill(3), _file[:filename_len_max - 4] + ".pdf"))
                else:
                    os.rename(_file, file_format.format(str(num).zfill(3), _file))
            num += 1

    def gen_excel_file(self, file_path, excel_file_name, filename_len_max):
        '''Generates excel file with new titles'''
        workbook = xlsxwriter.Workbook(excel_file_name)
        worksheet = workbook.add_worksheet()
        counter = 1

        for _file in os.listdir(file_path):
            path, ext = os.path.splitext(_file)
            if ext == '.pdf':
                worksheet.write('A' + str(counter), _file[:filename_len_max])
                counter += 1
        workbook.close()

    #Rename to copy an folder?
    def copy_job_folder(self, job_path, new_folder_name):
        '''Copies the specified job folder'''
        try:
            shutil.copytree(job_path, new_folder_name)
        except:
            print("Unable to copy template folder from source. Check your connection to network drives and try again...")
            print(self.fatal_error)


    def copy_files_to(self, destination_path):
        '''Copies files to new location'''
        livecount_path = destination_path
        try:
            os.chdir(livecount_path)
        except:
            print("Unable to navigate files. Check your connection to your network drive...")
            print(self.fatal_error)
        for _file in os.listdir(self.root_path):
            path, ext = os.path.splitext(_file)
            if ext == '.pdf':
                file_to_copy = os.path.join(self.root_path, _file)
                shutil.copy(file_to_copy, _file)

    def clean_up_meta(self, cleanup_path):
        '''Cleans up generated files that arent needed'''
        content = os.walk(cleanup_path)
        for data in content:
            path, folder, files = data
            for _file in files:
                if _file[:2] == '~$' or _file == 'Thumbs.db':
                    os.remove(os.path.join(path, _file))

    def clean_up_old(self, cleanup_path):
        '''Deletes pdf files from given path'''
        os.chdir(cleanup_path)
        for _file in os.listdir(cleanup_path):
            path, ext = os.path.splitext(_file)
            if ext == '.pdf':
                os.remove(_file)

    #All console statements should be in own file? not part of the class?
    def console_statement(self):
        '''Prints out statement to console while programs runs'''
        print(self.waiting_logo)
        print("\nWorking some very ipmressive estimating magic...")
        print("...might take a few seconds")

    def processing_statement(self):
        self.clear_terminal()
        self.console_statement()

    def exit_program(self):
        print("     Press ENTER to exit Tilley Renamer...")
        if input():
            os.system('exit')


    def organize(self):
        user_input = self.get_user_input()
        if user_input == '':
            self.processing_statement()
            self.gen_excel_file(self.root_path, self.excel_file_name, self.filename_len_max)
            #Build "Livecount Files" below as varaible- abstract it
            os.makedirs('Livecount Files')
            self.copy_files_to(self.root_path + '\Livecount Files')
        else:
            self.processing_statement()
            self.rename_files_numbered(self.root_path, self.filename_len_max, self.user_input)
            self.gen_excel_file(self.root_path, self.excel_file_name, self.filename_len_max)
            self.copy_job_folder(self.job_path, self.job_folder_name)
            self.copy_files_to(self.destination_path)

        self.clean_up_old(self.root_path)
        self.clean_up_meta(self.root_path)
        self.clear_terminal()
        print(self.success_logo)
        self.exit_program()
