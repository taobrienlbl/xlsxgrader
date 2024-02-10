import pandas as pd
from pathlib import Path

def parse_canvas_csv(csv_file_path : str | Path):
    """ Parse a CSV file exported from Canvas and return a cleaned-up DataFrame. """
    # load the CSV file; index is student name
    data = pd.read_csv(csv_file_path, index_col='name')

    # determine the number of questions
    # the columns are formatted as follows: id, section, section_id, submitted, attempt, question1, score1, question2, score2, question3, score3, ..., n correct, n incorrect, score
    # we can use the number of columns to determine the number of questions
    informational_columns_start = ['id', 'section', 'section_id', 'submitted', 'attempt']
    informational_columns_end = ['n correct', 'n incorrect', 'score']
    informational_columns = informational_columns_start + informational_columns_end
    num_informational_columns = len(informational_columns)
    num_informational_columns_start = len(informational_columns_start)
    num_informational_columns_end = len(informational_columns_end)
    num_columns = len(data.columns)
    num_questions = (num_columns - num_informational_columns) // 2

    # check that the number of columns is consistent with the number of questions
    if num_columns != num_informational_columns + 2 * num_questions:
        raise RuntimeError(f"Unexpected number of columns: {num_columns}")

    # extract the response columns
    response_columns = data.columns[num_informational_columns_start:-num_informational_columns_end:2]

    # extract the grade columns
    grade_columns = data.columns[num_informational_columns_start+1:-num_informational_columns_end:2]

    # extract the questions and the answers
    responses = data[response_columns]
    grades = data[grade_columns]

    # determine which questions need to be graded by checking if all the scores are 0
    grade_columns_that_sum_to_zero = data[grade_columns].sum() == 0

    # Extract the question text from each response column name
    question_text = [column.split(': ')[1] for column in responses.columns]

    # Extract the maximum score for each question from the grade column names
    # note that Pandas automatically appends a .1, .2, etc. to the column names if there are duplicates
    # so we need to extract the unique question numbers and then find the corresponding maximum score
    grade_column_names = grades.columns
    # check if there is more than one decimal point in the column names, and remove everything following the second decimal point
    maximum_grades = [int(column.split('.', 2)[0]) for column in grade_column_names]

    # generate new column names as Question 1, Question 2, etc.
    new_column_names = [f'Question {i+1}' for i in range(len(question_text))]

    # rename the columns for responses and grades
    responses.columns = new_column_names
    grades.columns = new_column_names

    # get the columns of the questions to grade (this is done after renaming the columns to Question 1, Question 2, etc.)
    responses_to_grade_bool = grade_columns_that_sum_to_zero.values
    responses_to_grade_columns = responses.columns[responses_to_grade_bool]

    # make a pandas dataframe for the maximum grades
    max_grades_df = pd.DataFrame(maximum_grades, index=new_column_names, columns=['Max Grade']).T

    # make a pandas dataframe for the question text
    question_text_df = pd.DataFrame(question_text, index=new_column_names, columns=['Question Text']).T

    # collect the question data into a dictionary
    question_data = {
        'question_text': question_text_df,
        'responses': responses,
        'grades': grades,
        'responses_to_grade': responses_to_grade_columns,
        'max_grades': max_grades_df
    }

    return question_data
