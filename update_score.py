"""
Author: Osama Bin Azam
Date: 06/01/2023
Description: This script is used to update the student scores in the Question file and export correspond csv file with updated scores.
"""

import sys
import subprocess
import platform
import pkg_resources
import pandas as pd
from time import sleep
import re
import os


# Clear the console
def clear_console():
    # Clear the console
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')

# This function is used to check and install the missing dependencies
import sys
import subprocess
import platform
import pkg_resources

def check_and_install_dependencies():
    # Check if Python is installed
    if platform.system() not in {'Windows', 'Linux', 'Darwin'}:
        print("Unsupported operating system. Please install Python manually.")
        sys.exit(1)

    # Check for required packages
    required_packages = {'pandas', 'numpy', 'openpyxl', 'xlrd'}

    # Determine currently installed packages
    installed_packages = {pkg.key for pkg in pkg_resources.working_set}
    missing_packages = required_packages - installed_packages

    # Install missing packages
    if missing_packages:
        print("Installing missing packages...")
        python = sys.executable
        subprocess.check_call([python, '-m', 'pip', 'install', *missing_packages], stdout=subprocess.DEVNULL)
    else:
        print("Python and required packages are already installed.")

# Set pandas display options to show complete data in console
def set_display_options():
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', None)


# This function is used to read the excel file and return the pandas ExcelFile object
def read_excel_file(file_path):
    # Read an Excel file and return a pandas ExcelFile object
    try:
        excel_file = pd.ExcelFile(file_path)
        return excel_file
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return None


# This function is used to process the excel file and return the dictionary of data
def process_excel_file(excel_file, skiprows=0, nrows=None):
    # Process the Excel file, reading each sheet and converting to a dictionary
    if excel_file is None:
        return {}
    

    sheet_data = {}
    sheets = excel_file.sheet_names
    if not nrows:
        # Loop through each sheet in the Excel file
        for sheet in sheets:
            # print(f"\nProcessing Sheet: {sheet}")

            # Read each sheet into a DataFrame
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet, skiprows=skiprows)

                # Convert the DataFrame to a list of dictionaries
                data_list = df.to_dict(orient='records')
                sheet_data[sheet] = data_list
            except Exception as e:
                print(f"Error processing sheet {sheet}: {e}")
    else:
        # Loop through each sheet in the Excel file
        for sheet in sheets:
            # print(f"\nProcessing Sheet: {sheet}")

            # Read each sheet into a DataFrame
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet, skiprows=0, nrows=nrows)

                # Convert the DataFrame to a list of dictionaries
                data_list = df.to_dict(orient='records')
                sheet_data[sheet] = data_list
            except Exception as e:
                print(f"Error processing sheet {sheet}: {e}")
    
    return sheet_data

# This  function help us to get the  current points of each  student
def getCurrentPoints(data):
    currentPoints = 0.0
    for row in data:
        if row['Is Correct'] == 'Correct':
            currentPoints += float(row['Points Received'])
    return currentPoints

# This  functions is used to calculate the total points of each student
def getTotalPoints(data):
    totalPoints = 0.0
    for row in data:
        totalPoints += float(row['Points Possible'])
    return totalPoints


# This function is used to generate the data for campus interface
def  generate_data_for_campus_interface(dqad_data , data_to_export):
    
    data = []
    # Replace 'your_file.csv' with the path to your CSV file
    file_path = os.path.join(os.getcwd(),'assets/2023-12-22T1733_Grades-APPLIED_PHARMACOLOGY__Florkey.csv')

    # Read  the CSV file
    df = pd.read_csv(file_path)

    csv_data  =  df.to_dict(orient='records')
    if not df.empty:
        # iterate through each sheet of the dqad file
        data_to_export.append(list(csv_data[0].values()))
        counter = 1;
        for sheet in dqad_data:
            if sheet  == csv_data[counter]['Student']:
                csv_data[counter]['Current Points'] = getCurrentPoints(dqad_data[sheet])
                csv_data[counter]['Final Points'] = getTotalPoints(dqad_data[sheet])

                csv_data[counter]['ID'] = int(csv_data[counter]['ID'])
                csv_data[counter]['SIS User ID'] = int(csv_data[counter]['SIS User ID'])
                csv_data[counter]['SIS Login ID'] = int(csv_data[counter]['SIS Login ID'])

                data_to_export.append(list(csv_data[counter].values()))
                counter += 1

        # Export the data to a CSV file
        return data_to_export
    else:
        return data_to_export

# This function is used to export the data to csv file
def export_data_to_csv(dqad):
    columns = [
    "Student"," ID", "SIS User ID", "SIS Login ID", "Section",
    "Assignment Current Points", "Assignments Final Points",
    "Assignments Current Score", "Assignments Unposted Current Score",
    "Assignments Final Score", "Assignments Unposted Final Score",
    "Current Points", "Current Score",
    "Unposted Current Score", "Final Score",
    "Unposted Final Score"
    ]

    data = []
    data.append(columns)
    data_to_export =generate_data_for_campus_interface(dqad, data)

    if not data_to_export:
        print("No data to export.")
        return
    
    
    df = pd.DataFrame(data_to_export)
    df.to_csv(os.path.join(os.getcwd(),'output/2023-12-22T1733_Grades-APPLIED_PHARMACOLOGY__Florkey.csv'), index=False, header=False)
   
    clear_console()

# This function is used to update excel files
def export_to_excel(data_dict, filename, student_info):
    """
    Export multiple DataFrames to an Excel file, each DataFrame in a separate sheet.

    :@param data_dict: Dictionary of DataFrames to export, with sheet names as keys
    :@param filename: Name of the Excel file to create
    """

    sheets = list(data_dict.keys())


    
    try:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            
            for sheet in sheets:
                data_frame = pd.DataFrame((data_dict[sheet]))
                if not data_frame.empty:
                    data_frame.to_excel(writer, sheet_name=sheet, index=False)
                else:
                    print(f"No data to export for sheet {sheet}.")
        print(f"File '{filename}' has been created successfully with multiple sheets.")
    except Exception as e:
        print(f"Error in exporting to Excel: {e}")


# Enter Questions for Credit
def enter_questions_for_credit():
    clear_console()
    questions_for_credit = {}

    print("Please follow the prompts to enter the questions and their corresponding details.")
    print("Type 'done' when you have finished entering all questions.\n")

    while True:
        try:
            question = input("Enter question number (e.g., 1, 2, 3) or type 'done' to finish: ")
            
            if question.lower() == 'done':
                break

            if not question.isdigit():
                raise ValueError("Invalid question number. Please enter a numeric value.")

            correct_answers_input = input(f"Enter correct answers for question {question} (e.g., Answer-one,Answer-2,Answer-3) separated by commas: ")
            if not correct_answers_input.replace(",", "").isalnum():
                raise ValueError("Invalid format for answers. Please enter letters separated by commas.")

            correct_answers = [answer.strip().upper() for answer in correct_answers_input.split(',')]

            points_input = input(f"Enter point value for question {question} (e.g., 5, 2.5): ")
            points = float(points_input)

            questions_for_credit[question] = {
                'correct_answers': correct_answers,
                'points': points
            }

            # print(f"\nQuestion {question} added with points: {points} and correct answers: {correct_answers}")

        except ValueError as e:
            print(f"Error: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    while True:
        print("\nReview the entered questions:")
        for q, details in questions_for_credit.items():
            print(f"Question {q}: Correct Answers: {details['correct_answers']}, Points: {details['points']}")

        review = input("\nAre all questions correct? (yes/no): ").lower()
        if review == 'yes':
            break
        elif review == 'no':
            question_to_edit = input("Enter the number of the question you want to edit: ")
            if question_to_edit in questions_for_credit:
                print(f"Editing question {question_to_edit}:")
                correct_answers_input = input("Enter correct answers (e.g., A,C,E) separated by commas: ")
                correct_answers = [answer.strip().upper() for answer in correct_answers_input.split(',')]
                points = float(input("Enter point value: "))

                questions_for_credit[question_to_edit] = {
                    'correct_answers': correct_answers,
                    'points': points
                }
            else:
                print("Invalid question number. Please try again.")
        else:
            print("Please answer 'yes' or 'no'.")
    clear_console()
    return questions_for_credit


# This function is used to handle multiple answers for a single question for comaprsion
def split_answers_and_keys(answer_string):

    # Check if the answer_string is a string
    if not isinstance(answer_string, str):
        return answer_string
    # Define the regex pattern for splitting by number and parenthesis (e.g., "1) ")
    number_pattern = r'\d+\)\s'
    # Define the pattern for splitting by semicolon
    semicolon_pattern = r';\s*'

    # Check if the string contains the number pattern
    if re.search(number_pattern, answer_string):
        # Split the string using the number pattern
        answers = list(filter(None, re.split(number_pattern, answer_string)))
    elif ';' in answer_string:
        # Split the string using the semicolon pattern
        answers = re.split(semicolon_pattern, answer_string)
    else:
        # If no patterns are found, return the string as a single-element list
        answers = [answer_string]

    return answers

# Update scores
def update_student_scores(question_file, questions_for_credit):

    if not questions_for_credit:
        print("No questions for credit have been entered. Updating student scores with existing data.")
        
    
    if not question_file:
        print("No data to update.")
        return -2
    

    for sheet in question_file:
        for row in question_file[sheet]:
            # converting answers to set 
            row['Student Answer'] = split_answers_and_keys(row['Student Answer'])
            row['Key'] = split_answers_and_keys(row['Key'])

            # Check if the question number is in the questions_for_credit dictionary
            if str(row['Question Number']) in questions_for_credit:
                question_info = questions_for_credit[str(row['Question Number'])]
                student_answers = set(row['Student Answer'])
                key_answers = set(question_info['correct_answers'])
                points_per_answer = float(question_info['points']) / len(key_answers)

                
                # Match student answers with  keys  by looping through each answer with key.
                matched_answers = []

                # Remove white spaces from student answers and key answers
                student_answers = [answer.strip() for answer in student_answers]
                key_answers = [answer.strip() for answer in key_answers]

                # Loop through each student answer and check if it is in the key answers
                for answer in student_answers:
                    if answer in key_answers:
                        matched_answers.append(answer)
                total_points = round(len(matched_answers) * points_per_answer, 2)

                row['Points Received'] = total_points
                row['Is Correct'] = 'Partial' if total_points > 0  and total_points <int(row['Points Possible']) else row['Is Correct']
                row['Points Possible'] = question_info['points']
                row['Key'] = question_info['correct_answers']

            else:
               
                student_answers = set(row['Student Answer'])
                key_answers = set(row['Key'])
                points_per_answer = float(row['Points Possible']) / len(key_answers)

                matched_answers = []

                # Remove white spaces from student answers and key answers
                student_answers = [answer.strip() for answer in student_answers]
                key_answers = [answer.strip() for answer in key_answers]

                # Loop through each student answer and check if it is in the key answers
                for answer in student_answers:
                    if answer in key_answers:
                        matched_answers.append(answer)
                total_points = round(len(matched_answers) * points_per_answer, 2)

                row['Points Received'] = total_points
                row['Is Correct'] = 'Partial' if total_points > 0 and total_points <int(row['Points Possible']) else row['Is Correct']

    clear_console()
    return 0

# This function is used to print the row
def print_question_analysis(data, questions_for_credit):
    if not data:
        print("No data to display.")
        return 
    
    if not questions_for_credit:
        print("No questions for credit have been entered.")
        return 
    
    sheets = list(data.keys())
    
    
    
    for sheet in sheets:
        print(f"\n\nStudent Name: {sheet}")
        for row in data[sheet]:
            if str(row['Question Number']) in questions_for_credit:
                print("Question No: ", row['Question Number'])
                print("Answers: ", row['Student Answer'])
                print("Keys : ", row['Key'])
                print("Points Received: ", row['Points Received'])
                print("Points Possible: ", row['Points Possible'])
                print("Is Correct: ", row['Is Correct'])
                print("\n\n")
    

        # Ask the user if they want to see the next sheet
        next_sheet = input("\n(Type 'yes' to jump on next student detail or press Enter to go back to menu: ")
        
        clear_console()
        # If the user enters anything other than 'yes', break out of the loop
        if next_sheet.lower() != 'yes':
            break

# This function is used to display Menu  to the user 
def menu(dqad_data, ace_data, student_info):

    print("\n\nEverything looks good. Please  select the  option from the  menu below")
    questions_for_credit= {}
    while True:
        print("Menu:")

        print("1. Enter Questions for Credit")
        print("2. Update Student Scores")
        print("3. Create Formatted File for Campus Interface")
        print("4. Print Question Analysis")
        print("5. Exit")

        choice = input("Enter your choice (1-5): ")
        if choice == '1':
            questions_for_credit = enter_questions_for_credit()
            print("Questions for credit have been entered.")
           
        elif choice == '2':
            # Placeholder for function to update student scores
            print("Updating student scores...")
            response =  update_student_scores(dqad_data, questions_for_credit)
            if response == -1:
                print("No questions for credit have been entered. Updating student scores with existing data.")
            elif response == -2:
                print("No data to update.")
            else:
                export_to_excel(dqad_data, os.path.join(os.getcwd(),'output/DetailQuestionAnalysis (22).xlsx'), student_info)
                print("Student scores have been updated.")
                
        
        elif choice == '3':
            # Placeholder for function to create formatted file for campus interface
            print("Exporting data to CSV file...")
            export_data_to_csv(dqad_data)
            print("Data exported to CSV file.")
        
        elif choice == '4':
            print("")
            print_question_analysis(dqad_data, questions_for_credit)
        
        elif choice == '5':
            print("Exiting the program.")
            break
        else:
            clear_console()
            print("Invalid choice. Please try again.")


# This function is used to handle pre processing of libraries and  data and then call the menu function
def main():

    print("Welcome to Script!\n")
    print("Checking for missing dependencies...")
    # Check for and install any missing dependencies
    try:
        check_and_install_dependencies()
    except Exception as e:
        print(f"Error checking for missing dependencies: {e}")
        return

    # Set display options for pandas
    set_display_options()

    print("Loading Excel files...")
    try:
        # File paths for the Excel files to process
        dqad_file = os.path.join(os.getcwd(),'assets/DetailQuestionAnalysis (22).xlsx')
        ace_file = os.path.join(os.getcwd(),'assets/AssessmentID Proctored Assessment CSV Export (11).xls')
    except Exception as e:
        print(f"Error loading Excel files!. Please Check the file path or files in the asset folder: ")
        return

    print("Processing Excel files...")
    try:
        # Load and process the Excel files
        dqad_excel = read_excel_file(dqad_file)
        ace_excel = read_excel_file(ace_file)

        # Process the data with specified skiprows
        dqad_data = process_excel_file(dqad_excel, skiprows=5)
        ace_data = process_excel_file(ace_excel)
    except Exception as e:
        print(f"Error loading Excel files!, Please check the  format of the excel files:")
        return
    
    student_info = process_excel_file(dqad_excel, nrows=5)
    

    # export_to_excel(dqad_data, os.path.join(os.getcwd(),'output/DetailQuestionAnalysis (22).xlsx'))
    # Display the menu
    # clear_console()
    menu(dqad_data, ace_data, student_info)


if __name__ == "__main__":
    main()
