import os
import shutil
from typing import Counter
from zipfile import ZipFile
import zipfile

zip_file_name = "submissions.zip"
current_chapter = ""
excel_file_path = "V:\Work Study\CMPR112\CMPR112 - Chapter NA - Assignment_Report.xlsx"
current_dir = os.getcwd()
destination_path = "V:\Work Study\CMPR112\\"
working_path = ""


def create_next_chapter_directory_from_previous():
    try:
        max_num = max([chapter[-1:] for chapter in os.listdir() if chapter[:7] == "Chapter" ])
    except:
        max_num = 0;
    new_folder_name = "Chapter " + str(int(max_num)+1)
    os.mkdir(new_folder_name)
    
    global current_chapter; current_chapter = new_folder_name
    global working_path; working_path = destination_path + current_chapter + "\\"
    current_chapter = new_folder_name
    return new_folder_name


def moving_submission_zip_file_to_working_directory():
    try:
        shutil.move(os.path.join(current_dir,zip_file_name), os.path.join(current_dir,current_chapter))
        print("Moving the submissions.zip to the " + current_chapter + " directory.")
    except:
        print("Submission zip file is missing. Please check. Quitting")
        quit()

def check_if_excel_file_exist(current_dir):
    execl_file_name = "xlsx"
    for file_name in current_dir:
        if (file_name[-4:] == execl_file_name):
            os.remove(destination_path + current_chapter + "\\" + file_name)
            return "CMPR112 - Chapter NA - Assignment_Report.xlsx"
    return "CMPR112 - Chapter NA - Assignment_Report.xlsx"
        
def check_if_zip_file_exist():
    current_working_dir = os.listdir(os.path.join(os.getcwd(),current_chapter))
    for file_name in current_working_dir:
        if (file_name == zip_file_name):
            print("Zip file exist in the current directory, moving to the next step.")
            return
    print("Zip file is not exist in the current directory. Quitting")
    quit()

def copying_excel_report_from_main_directory():
    # Check if the excel file exist or not in the current directory
    excel_file_name = check_if_excel_file_exist(current_dir)
    try:
        shutil.copy2(excel_file_path,os.path.join(destination_path,current_chapter))
    except:
        print("Excel file is missing. Please check. Quitting.")
        quit()

    excel_file_new_name = "CMPR112 - " + current_chapter + " - Assignment_Report.xlsx"
    os.rename(working_path + excel_file_name, working_path + excel_file_new_name)

def createDirectoryIfNotExist(fileName):
    if not (os.path.isdir(fileName)):
        os.mkdir(fileName)

def exacting_folder_from_zips(list_of_all_projects):
    for i in list_of_all_projects:
        with ZipFile(working_path + i,'r') as zipObj:
            createDirectoryIfNotExist(working_path + i[:-3])
            zipObj.extractall(working_path + i[:-3])
            print(i + " is unzipped")

def extracting_submission_zip_file():
    with ZipFile(os.path.join(working_path,zip_file_name),'r') as zipObj:
        zipObj.extractall(working_path)
        print("submissions zip is extracted.")

    current_dir = os.listdir(os.path.join(os.getcwd(),current_chapter));

    list_of_all_projects = [project for project in current_dir if project[-3:] == "zip" and project != zip_file_name]   
    index = 1;

    for i,element in enumerate(list_of_all_projects):
        new_string = str(index) + " - " + element
        os.rename(working_path + element,working_path + new_string)
        print(element + " is renaming to " + new_string)
        list_of_all_projects[i] = new_string
        index += 1
    
    return list_of_all_projects

if __name__ == "__main__":
    create_next_chapter_directory_from_previous()
    
    moving_submission_zip_file_to_working_directory()

    copying_excel_report_from_main_directory()

    list_of_all_projects = extracting_submission_zip_file()

    exacting_folder_from_zips(list_of_all_projects)

    

    



