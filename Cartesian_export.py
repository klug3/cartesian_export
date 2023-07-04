import os
import pprint
from operator import itemgetter
import pandas as pd
import docx
from natsort import natsorted

# Define the extension of analyzed file
# Available options: "out" or "gjf"
FILE_EXT = "gjf"


def files_list():
    """
    Walk down through dirs searching all files meeting specified requirements.
    Pair file name and its root localization inside the tuple and add the tuple to list.
    :return: f - list of tuples containing file name and its localization
    """
    global FILE_EXT
    path = os.getcwd()
    f = []
    if FILE_EXT == "out":
        for root, dirs, files in os.walk(path):
            for file_name in files:
                if file_name[-3:] == "out":
                    path = os.path.join(root, file_name)
                    f.append((path, file_name))
    elif FILE_EXT == "gjf":
        for root, dirs, files in os.walk(path):
            for file_name in files:
                if file_name[-3:] == "gjf":
                    path = os.path.join(root, file_name)
                    f.append((path, file_name))
    elif FILE_EXT != "out" or FILE_EXT != "gjf":
        print("Wprowadzono błędne rozszerzenie pliku! Dostępne rozszerzenia: out i gjf.")
        print("Dalsze działanie programu nie ma sensu. Zmodyfikuj zmienną FILE_EXT (9 linia) i uruchom ponownie.")
        exit("Program zakończył pracę.")
    elif not f:
        print("W danym folderze oraz jego podfolderach nie znaleziono odpowiednich plików.")
        print("Dalsza praca programu nie ma sensu.")
        exit("Program kończy pracę.")
    else:
        print("Brawo! Tego błędu nie przewidziałem.")
        exit("Program kończy pracę.")


    # Sort the list of tuples by file_name
    f = natsorted(f, key=itemgetter(1))
    return f


def read_out_file(file_tuple):
    """
    :arg file_tuple - tuple containing (path, file_name)
    Open, read and close .txt file
    """
    # Unpack data from tuple
    path, file_name = file_tuple
    # Read gjf or out file
    with open(path, "r") as file:
        lines = file.readlines()
        # When FILE_EXT is set to "out" search for final coordinates table
        if FILE_EXT == "out":
            # Create dictionary for isolated data
            data = {'Center Number': [],
                    'Atomic Number': [],
                    'Atomic Type': [],
                    'Coordinate X': [],
                    'Coordinate Y': [],
                    'Coordinate Z': []}
            # Set beginning and end of the table
            for nr, value in enumerate(lines):
                if "Standard orientation:" in value:
                    start_pos = nr + 5
                elif value.strip() == "---------------------------------------------------------------------":
                    end_pos = nr
            # Gather data from file
            for entry in lines[start_pos:end_pos]:
                centr_num, at_num, at_type, co_x, co_y, co_z = entry.split()
                data['Center Number'].append(centr_num)
                data['Atomic Number'].append(at_num)
                data['Atomic Type'].append(at_type)
                data['Coordinate X'].append(co_x)
                data['Coordinate Y'].append(co_y)
                data['Coordinate Z'].append(co_z)
        # When FILE_EXT is set to "gjf" read the coordinates table from the beginning of the file
        else:
            # Create dictionary for isolated data
            data = {'Atom Symbol': [],
                    'Coordinate X': [],
                    'Coordinate Y': [],
                    'Coordinate Z': []}
            # Set beginning and end of the table
            start_pos = 5   # in every gjf file table starts in 6th line
            for nr, value in enumerate(lines):
                if not value.strip() and nr > start_pos:
                    end_pos = nr
                    break
            # Gather data from file
            for entry in lines[start_pos:end_pos]:
                atom_sym, co_x, co_y, co_z = entry.split()
                data['Atom Symbol'].append(atom_sym)
                data['Coordinate X'].append(co_x)
                data['Coordinate Y'].append(co_y)
                data['Coordinate Z'].append(co_z)
        # Create dataframe out of isolated data
        df = pd.DataFrame(data)
        # Format filename removing .out and .gjf
        while file_name[-4:] == ".out" or file_name[-4:] == ".gjf":
            file_name = file_name[:-4]

        return file_name, df


def fill_docx_file():
    """
    Open blank docx file, create table, fill it with data and saves the file.
    :return: nothing
    """
    # Create list of files
    file_list = files_list()

    # Open docx file and create table
    doc = docx.Document('Cartesians.docx')
    t = doc.add_table(len(file_list) + 1, 2)

    # Add the header rows.
    t.cell(0, 0).text = "Conformer"
    t.cell(0, 1).text = "Coordinates"

    # Add the rest of the data
    for nr, tuple in enumerate(file_list):
        file_name, data_frame = read_out_file(tuple)
        dane = data_frame.to_string(index=False, header=False)
        t.cell(nr + 1, 1).text = dane
        t.cell(nr + 1, 0).text = file_name

    # save the doc
    doc.save('Cartesians.docx')


# Main
input("Wio!")
print()
print("Zaczynam robotę...")
fill_docx_file()
print()
print("Zrobione!!!!")

input()
