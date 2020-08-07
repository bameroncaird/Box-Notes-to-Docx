import os
from boxnotes2html import BoxNote

def get_file_name(file_path):
    return os.path.splitext(file_path)[0]

def get_small_name(file_path):
    start = len(file_path) - 1
    if file_path[start] == '/':
         start -= 1
    while start >= 0 and file_path[start] != '/':
        start -= 1
    start += 1
    output = ""
    while start <= (len(file_path) - 1):
        output += file_path[start]
        start += 1
    return output

class BoxNoteConverter:

    # Instance Variables:
    root_folder = None  # Specifies the starting folder as a file path.
    removes_box_notes = None  # Whether or not to remove the .boxnote files after conversion
    removes_html_files = None  # Whether or not to remove temporary .html files after conversion

    # Constructors:
    def __init__(self):
        self.root_folder = ""
        self.removes_box_notes = False
        self.removes_html_files = True
    def __init__(self, root_folder, removes_box_notes, removes_html_files):
        self.root_folder = root_folder
        self.removes_box_notes = removes_box_notes
        self.removes_html_files = removes_html_files


    # Methods:

    def convert_box_to_doc_b(self, file_path):

        if not os.path.exists(file_path):
            print("The .boxnote file does not exist.")
            return

        base_file_name = get_file_name(file_path)
        file_destination_directory = "/Users/bameroncaird/Box/Testing/output_testing/"

        html_file_name = base_file_name + ".html"
        docx_file_name = file_destination_directory + get_small_name(base_file_name) + ".docx"

        if os.path.exists(docx_file_name):
            print("The .docx file already exists...")
            return

        # THIS IS THE LINE THAT IS CAUSING ERRORS
        note = BoxNote.from_file(file_path)

        # First, convert to html
        note_as_html = note.as_html()
        tmp_html = open(html_file_name, "w")
        tmp_html.write(note_as_html)
        tmp_html.close()

        # Next, convert to docx with the pandoc command
        # An old command that had some issues with images:
        # command = "pandoc -f html -t docx -o " + docx_file_name + " " + html_file_name

        command = "pandoc -s " + html_file_name + " -o " + docx_file_name
        os.system(command)

        if self.removes_html_files:  # Remove the HTML file you created
            os.remove(html_file_name)

        # if self.removes_box_notes:
        #     # Remove the original box note file

        print("Original file path: " + file_path)
        print("Base file name: " + base_file_name)
        print("HTML file name: " + html_file_name)
        print("docx file name: " + docx_file_name)
        print("pandoc command: " + command)

    # Converts a single .boxnote file to a .docx file - having some issues
    def convert_box_to_doc(self, file_path):

        if not os.path.exists(file_path):
            print("The .boxnote file does not exist.")
            return

        base_file_name = get_file_name(file_path)

        html_file_name = base_file_name + ".html"
        docx_file_name = base_file_name + ".docx"
        
        if os.path.exists(docx_file_name):
            print("The .docx file already exists...")
            return

        # THIS IS THE LINE THAT IS CAUSING ERRORS
        note = BoxNote.from_file(file_path)

        # First, convert to html
        note_as_html = note.as_html()
        tmp_html = open(html_file_name, "w")
        tmp_html.write(note_as_html)
        tmp_html.close()

        # Next, convert to docx with the pandoc command
        # An old command that had some issues with images:
        # command = "pandoc -f html -t docx -o " + docx_file_name + " " + html_file_name

        command = "pandoc -s " + html_file_name + " -o " + docx_file_name
        os.system(command)

        if self.removes_html_files:  # Remove the HTML file you created
            os.remove(html_file_name)

        # if self.removes_box_notes:
        #     # Remove the original box note file

        # print("Original file path: " + file_path)
        # print("Base file name: " + base_file_name)
        # print("HTML file name: " + html_file_name)
        # print("docx file name: " + docx_file_name)
        # print("pandoc command: " + command)


    # This recursive method loops through the root directory and, if specified, the sub-folders, converting .boxnote files to .docx files as it goes along.
    def convert_files(self, directory_path):

        num_box_notes_before = self.get_number_of_box_notes(directory_path)
        num_docs_before = self.get_number_of_docs(directory_path)

        for root, dirs, files in os.walk(directory_path):
            for file in files:
                if file.endswith('.boxnote'):
                    print(os.path.join(root, file))
                    self.convert_box_to_doc(os.path.join(root, file))

        num_docs_after = self.get_number_of_docs(directory_path)
        print("of " + num_box_notes_before + " .boxnote files, " + (num_docs_after - num_docs_before) + " were converted to .docx format.")
                    
    def get_number_of_box_notes(self, file_path) -> int:
        count = 0
        for root, dirs, files in os.walk(file_path):
            for file in files:
                if file.endswith('.boxnote'):
                    count += 1
        return count

    def get_number_of_docs(self, file_path) -> int:
        count = 0
        for root, dirs, files in os.walk(file_path):
            for file in files:
                if file.endswith('.docx'):
                    file_name = os.path.join(root, file)
                    # print(file_name)
                    count += 1
        return count


    # This method will convert files from .boxnote to .docx starting at whatever the root directory is set to.
    def convert(self):
        self.convert_files(self.root_folder)


    # Methods for setting properties of the class:
    def set_root_folder(self, root_folder):
        self.root_folder = root_folder
    def set_html_removal(self, removes_html_files):
        self.removes_html_files = removes_html_files
    def set_box_note_removal(self, removes_box_notes):
        self.removes_box_notes = removes_box_notes








