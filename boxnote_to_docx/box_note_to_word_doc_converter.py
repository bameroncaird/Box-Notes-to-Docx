import os
from boxnotes2html import BoxNote

# Returns the number of box notes in a folder and any sub-folders that the folder may have.
# Takes as input a string that is a path to a folder in the filesystem.
def get_number_of_box_notes_in_folder(path_to_folder) -> int:
    count = 0
    for root, dirs, files in os.walk(path_to_folder):
        for file in files:
            if file.endswith('.boxnote'):
                count += 1
    return count

# Prints all the box note file names in a folder and any sub-folders that the folder may have.
# Takes as input a string that is a path to a folder in the filesystem.
def print_all_box_notes(path_to_folder):
    num = get_number_of_box_notes_in_folder(path_to_folder)
    print("Printing " + str(num) + " Box Notes...")
    for root, dirs, files in os.walk(path_to_folder):
        for file in files:
            if file.endswith('.boxnote'):
                print(os.path.join(root, file))

# Prints out all the Box notes in a folder and any sub-folders in a more organized way by separating the box notes by folder.
def print_box_notes_separated_by_folder(path_to_folder):
    print("Printing Box Notes separated by folder...\n")
    count = 0
    count += print_box_notes_in_single_folder(path_to_folder)
    for root, dirs, files in os.walk(path_to_folder):
        for directory in dirs:
            count += print_box_notes_in_single_folder(os.path.join(root, directory))
    print("Printed " + str(count) + " total Box Notes.")

# Prints all the Box Notes in a single folder and returns the number of Box Notes that were printed.
def print_box_notes_in_single_folder(path_to_folder) -> int:
    print("Folder: " + path_to_folder + "\nBox Notes: ")
    count = 0
    for item in os.listdir(path_to_folder):
        if item.endswith(".boxnote"):
            count += 1
            print(item)
    if count == 0:
        print("There were no boxnote files in the folder.\n")
    else:
        print("Printed " + str(count) + " Box Notes in folder " + path_to_folder + ".\n")
    return count


# Returns the number of word docs in a folder and any sub-folders that the folder may have.
# Takes as input a string that is a path to a folder in the filesystem.
def get_number_of_word_docs_in_folder(path_to_folder) -> int:
    count = 0
    for root, dirs, files in os.walk(path_to_folder):
        for file in files:
            if file.endswith('.docx'):
                count += 1
    return count

# Prints all word doc file names in a folder and any sub-folders that the folder may have.
# Takes as input a string that is a path to a folder in the filesystem.
def print_all_word_docs(path_to_folder):
    num = get_number_of_word_docs_in_folder(path_to_folder)
    print("Printing " + str(num) + " Word Docs...")
    for root, dirs, files in os.walk(path_to_folder):
        for file in files:
            if file.endswith('.docx'):
                print(os.path.join(root, file))

# Returns the total number of pairs of box notes and word docs in the directory.
def get_number_of_pairs(path_to_folder):
    pairs = 0
    for root, dirs, files in os.walk(path_to_folder):
        for file in files:
            if file.endswith(".boxnote"):  # Check for a pair
                base_file_name = get_file_name(os.path.join(root, file))
                doc_file_name = base_file_name + ".docx"
                if os.path.exists(doc_file_name):
                    pairs += 1
    return pairs

# Prints out the box note file names that don't have corresponding word docs.
def print_lonely_box_notes(path_to_folder):
    print("\nLooking for lonely Box Notes...")
    lonely_notes = 0
    for root, dirs, files in os.walk(path_to_folder):
        for file in files:
            if file.endswith(".boxnote"):
                base_file_name = get_file_name(os.path.join(root, file))
                doc_file_name = base_file_name + ".docx"
                if not os.path.exists(doc_file_name):
                    print("Lonely note found: " + os.path.join(root, file))
                    lonely_notes += 1
    if lonely_notes:
        print("\nFound " + str(lonely_notes) + " lonely Box Notes.\n")
    else:
        print("\nFound no lonely Box Notes (all are paired).")

# The next two methods are for formatting for amanda.
def parse_lonely_note(path_to_file):
    return path_to_file.split("/", 3)[3]

def print_lonely_notes_for_amanda(path_to_folder):
    num = 1
    for root, dirs, files in os.walk(path_to_folder):
        for file in files:
            if file.endswith(".boxnote"):
                base_file_name = get_file_name(os.path.join(root, file))
                doc_file_name = base_file_name + ".docx"
                if not os.path.exists(doc_file_name):
                    lonely_note = os.path.join(root, file)
                    print(str(num) + ". " + parse_lonely_note(lonely_note) + "\n")
                    num += 1

# Returns a list of Box Notes without corresponding Word Docs.
def get_lonely_box_notes(path_to_folder):
    lonely_notes = []
    for root, dirs, files in os.walk(path_to_folder):
        for file in files:
            if file.endswith(".boxnote"):
                base_file_name = get_file_name(os.path.join(root, file))
                doc_file_name = base_file_name + ".docx"
                if not os.path.exists(doc_file_name):
                    lonely_notes.append(os.path.join(root, file))
    return lonely_notes


# Returns the name of a file without the extension by splitting the path to the file at the period.
# os.path.splitext() returns a tuple (name, extension) that splits the string at the period.  Read more: https://docs.python.org/2/library/os.path.html
def get_file_name(path_to_file) -> str:
    return os.path.splitext(path_to_file)[0]

# Returns the extension of a file without the name by splitting the path to the file at the period.
# If there is no period, the method returns an empty string.
# os.path.splitext(): https://docs.python.org/2/library/os.path.html
def get_file_extension(path_to_file) -> str:
    if '.' not in path_to_file:
        print("The extension will be empty...")
    return os.path.splitext(path_to_file)[1]

# Returns the name of the file with extension relative to the present working directory (pwd).
# str.split() returns a list of strings that are split by a specified character. Read more: https://www.w3schools.com/python/ref_string_split.asp
def get_file_name_in_pwd(path_to_file) -> str:
    split_list = path_to_file.split("/")
    return split_list[len(split_list) - 1]

# Returns the path to the parent directory of the current file.
def get_parent_directory_name(path_to_file) -> str:
    split_list = path_to_file.rsplit("/", 1)
    return split_list[0]

# On the command line, spaces need to be escaped before use, so every file will be the file name but have the spaces replaced with "\ "
def parse_file_name_for_cli(path_to_file) -> str:
    output_string = ""
    for character in path_to_file:
        if character == " ":
            output_string += "\ "
        elif character == "(":
            output_string += "\("
        elif character == ")":
            output_string += "\)"
        elif character == "&":
            output_string += "\&"
        elif character == "'":
            output_string += "\\'"
        elif character == " ":
            output_string += "\ "
        elif character == '"':
            output_string += '\\"'
        else:
            output_string += character
    return output_string


# "Converts" a single Box Note to a Word Doc with the same name and in the same directory by first converting to an HTML file.
# Returns True if the conversion was successful and False otherwise.
# Quotation marks are there because I am having some issues.  Read comments in method to locate issues.
def convert_single_box_note(path_to_file, removes_html_files, removes_box_notes) -> bool:

    # base_file_name is the full path with file name without the extension.
    # Example: /Users/bameroncaird/Box/Testing/random_note_1.boxnote -> /Users/bameroncaird/Box/Testing/random_note_1
    base_file_name = get_file_name(path_to_file)

    html_file_name = base_file_name + ".html"
    docx_file_name = base_file_name + ".docx"

    parsed_name = parse_file_name_for_cli(path_to_file)
    print(parsed_name)

    if os.path.exists(html_file_name):
        os.remove(html_file_name)

    if not os.path.exists(path_to_file):
        print("The path to the .boxnote file does not exist and so the file was not converted.")
        return False

    if os.path.exists(docx_file_name):
        print("There is already a .docx file with the same name in the directory, so the file was not converted out of unnecessity.")
        return False

    # The line after these comments "randomly" fails sometimes - storing the BoxNote from the file in a variable.
    # Quotation marks are there because there is a reason, I just don't know what it is and have not figured it out.
    # This line uses boxnotes2html, which is a Python package with a self-explanatory name. Read more: https://pypi.org/project/boxnotes2html/
    note = BoxNote.from_file(path_to_file)

    # The next four lines convert the BoxNote file to HTML.
    # as_html comes from boxnotes2html, which you can read more about above if you want.
    try:
        note_as_html = note.as_html()
    except:
        print("Could not convert the file to .docx format.")
        return False

    tmp_html = open(html_file_name, "w")
    tmp_html.write(note_as_html)
    tmp_html.close()

    # The next two lines execute a pandoc command in terminal to "convert" from HTML to Word Doc:
    # Read more about pandoc: https://pandoc.org/demos.html
    # Read more about os.system: https://www.geeksforgeeks.org/python-os-system-method/
    # THERE ARE ISSUES HERE!  Note - Issue #2 I believe has been resolved!
    #   1. If the HTML files have images, they won't be displayed in the resulting Word Doc.  No idea how to fix
    parsed_html_file_name = parse_file_name_for_cli(html_file_name)
    parsed_doc_file_name = parse_file_name_for_cli(docx_file_name)
    command = "pandoc -s " + parsed_html_file_name + " -o " + parsed_doc_file_name
    os.system(command)

    # os.remove just deletes a file from your file system, here I use it based on parameters. Read more: https://www.geeksforgeeks.org/python-os-remove-method/
    if removes_box_notes:
        os.remove(path_to_file)
    if removes_html_files:
        os.remove(html_file_name)

    return True

    # I put some print statements here to make sure my file names were correct:
    # print("Original file path: " + file_path)
    # print("Base file name: " + base_file_name)
    # print("HTML file name: " + html_file_name)
    # print("docx file name: " + docx_file_name)
    # print("pandoc command: " + command)

# Calls a method to convert a Box Note to a Word Doc for each Box Note in a specified folder, as well as any sub-folders that the folder may have.
# Calculates the number of successful conversions from the number of attempted conversions and prints both out to the console.
def convert_files(path_to_folder):

    prev_box_notes = get_number_of_box_notes_in_folder(path_to_folder)

    print("Trying to convert " + str(prev_box_notes) + " Box Notes to Word Doc format...")

    successful_conversions = 0
    for root, dirs, files in os.walk(path_to_folder):
        for file in files:
            if file.endswith(".boxnote"):
                if convert_single_box_note(os.path.join(root, file), True, False):  # If the method returns True
                    successful_conversions += 1
    print("Of " + str(prev_box_notes) + " attempted Box Note to Word Doc conversions, " + str(successful_conversions) + " were successful.\n")

def check_folder(path_to_folder):
    print("\nFolder: " + path_to_folder)
    num_box = get_number_of_box_notes_in_folder(path_to_folder)
    num_doc = get_number_of_word_docs_in_folder(path_to_folder)
    num_pairs = get_number_of_pairs(path_to_folder)
    print("\nBox Notes: " + str(num_box) + "\nWord Docs: " + str(num_doc) + "\nPairs: " + str(num_pairs))

    if num_box == num_pairs:
        print("All the files in folder " + path_to_folder + " should be converted - pairs are equal to the number of .boxnotes.")
    else:
        print("There were some files - " + str(num_box - num_pairs) + " - that were not converted to .docx format.")



root_path = "/Users/bameroncaird/Box/A-Team"

# specific_path = '/Users/bameroncaird/Box/A-Team/Web Projects/Archives/1. Why Hire a Tiger info on HMT - "CREAM of Crop" edit/Instructions: Rotating ads and new page content.boxnote'
# convert_single_box_note(specific_path, True, False)

# print_lonely_box_notes(root_path)

print_lonely_notes_for_amanda(root_path)

