# Import libraries
import pandas as pd

def extract_from_excel(filename, sheetname):
    # Open excel file
    df = pd.read_excel(filename, sheetname)

    # Extract all values that are strings
    info = [[value for value in row if type(value) == str] for row in df.values]

    # Extract the puzzle by getting all the strings with one letter
    puzzle = [[value for value in row if len(value) == 1] for row in info]
    
    # Remove all the empty rows from puzzle 
    puzzle = [row for row in puzzle if len(row) != 0]

    
    # Extracts the wordlist from the info by getting all the values their length is > 1 and flattens 2D array
    # Except for if it has "Oplossing" in the string, which is the solution words
    wordlist = sum([[value for value in row if len(value) > 1] for row in info], [])
    wordlist = [word for word in wordlist if not "oplossing" in word.lower()]
    
    return puzzle, wordlist