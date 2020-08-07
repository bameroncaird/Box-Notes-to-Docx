# The following lines append the docx library to the path and then save the file as a .docx.
# import sys
# sys.path.append('/Library/Frameworks/Python.framework/Versions/3.8/lib/python3.8/site-packages')
#
# import docx
# document = docx.Document()
# document.add_heading("Let's see how this works...")
# document.save('/Users/bameroncaird/Box/Testing/TEST.docx')
#
# from boxnotes2html import BoxNote
# import os
# Gets the corresponding .boxnote file from whatever directory.
# note = BoxNote.from_file("/Users/bameroncaird/Box/Testing/root_note_3.boxnote")

# - Some more functions from the boxnotes2html module.
# note.as_html() # returns an html string
# note.as_markdown() # returns a markdown string

# - The next 5 lines of code create a new .txt file, convert it to a .docx file, and save it in the same folder.
# textNote = note.as_text()  # returns raw text
# testFile = open("/Users/bameroncaird/Box/Testing/test_file_1.txt", "w")
# testFile.write(textNote)
# testFile.close()
# os.system("pandoc -s /Users/bameroncaird/Box/Testing/test_file_1.txt -o /Users/bameroncaird/Box/Testing/test_file_1.docx")

# pandoc -s MANUAL.txt -o example29.docx

# print(textNote)

import os
from boxnotes2html import BoxNote

# Converts from markdown to docx - this will not be used.
# def convert_to_docx_from_markdown(boxnote):
#     note_as_markdown = boxnote.as_markdown()
#     markdown_file = open("/Users/bameroncaird/Box/hello_md.md", "w")
#     markdown_file.write(note_as_markdown)
#     markdown_file.close()
#     os.system("pandoc -s -o /Users/bameroncaird/Box/hello_doc.docx /Users/bameroncaird/Box/hello_md.md")

#
# def get_file_name(file_path):
#     return os.path.splitext(file_path)[0]
#
# def get_directory_from_file(file_path):
#     return os.path.dirname(file_path)
#
# def convert_box_to_doc(file_path):
#     base_file_name = get_file_name(file_path)
#
#     html_file_name = base_file_name + ".html"
#     docx_file_name = base_file_name + ".docx"
#
#     note = BoxNote.from_file(file_path)
#
#     # First, convert to html
#     note_as_html = note.as_html()
#     tmp_html = open(html_file_name, "w")
#     tmp_html.write(note_as_html)
#     tmp_html.close()
#
#     # Next, convert to docx with the pandoc command
#     command = "pandoc -f html -t docx -o " + docx_file_name + " " + html_file_name
#     os.system(command)
#
#     print("Base file name: " + base_file_name)
#     print("HTML file name: " + html_file_name)
#     print("docx file name: " + docx_file_name)
#     print("pandoc command: " + command)
#
# convert_box_to_doc("/Users/bameroncaird/Box/TestBoxNote.boxnote")
#
#
# def convert_to_docx(boxnote):
#     note_as_html = boxnote.as_html()
#
#     # Create an HTML file from the boxnote:
#     html_file = open("/Users/bameroncaird/Box/summer_projects.html", "w")
#     html_file.write(note_as_html)
#     html_file.close()
#
#     # HTML to DOCX with pandoc.
#     os.system("pandoc -f html -t docx -o /Users/bameroncaird/Box/mj_docx.docx /Users/bameroncaird/Box/mj_html.html")

# note_to_convert = BoxNote.from_file("/Users/bameroncaird/Box/A-Team/Web Projects/Summer 2020 Project List.boxnote")
# convert_to_docx(note_to_convert)
#
#
# os.system("pandoc -f html -t docx -o /Users/bameroncaird/Box/mj_docx.docx /Users/bameroncaird/Box/index.html")

# note_to_convert = BoxNote.from_file("/Users/bameroncaird/Box/index.html")
# convert_to_docx(note_to_convert)
# convert_to_docx_from_markdown(note_to_convert)

# print(note_to_convert.as_html())

# for root, dirs, files in os.walk(root_directory):
#     for file in files:
#         print(os.path.join(root, file))
#     for directory in dirs:
#         print(os.path.join(root, directory))


# This code prints all of the box notes in a given directory:
# def print_all_box_notes(main_directory):
#     for root, dirs, files in os.walk(main_directory):
#         for file in files:
#             if file.endswith('.boxnote'):
#                 print(os.path.join(root, file))
#         for directory in dirs:
#             print_all_box_notes(directory)
#
# root_directory = "/Users/bameroncaird/Box"
# print_all_box_notes(root_directory)



# print(os.getcwd())

# from BoxNoteConverter import BoxNoteConverter
#
# root_path = "/Users/bameroncaird/Box/A-Team"
#
# converter = BoxNoteConverter(root_path, False, True)
# converter.set_html_removal(False)
#
# converter.convert()












