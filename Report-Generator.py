# Import re for regex functions
import re
# Import docx to work with .docx files.
from docx import Document

# Store file paths from input
template_path = input('Write the full path to the template file: ')
list_path = input('Write the full path to the list file: ')
to_replace = input('Write a word to replace: ')

if template_path.endswith('.docx') and list_path.endswith('.docx'):
    words = Document(list_path)
    
    # Loop through replacer arguments
    for replacement in words.paragraphs:
        doc = Document(template_path)
        print("\nСurrent replacement text: " + replacement.text)
        # initialize the number of occurrences of this word to 0
        occurrences = {to_replace: 0}
        
        # Loop through paragraphs
        for paragraph in doc.paragraphs:
            # Loop through runs (style spans)
            for run in paragraph.runs:
                # if there is text on this run, replace it
                if run.text:
                    # get the replacement text
                    replaced_text = re.sub(to_replace, replacement.text, run.text, 999)
                    if replaced_text != run.text:
                        # if the replaced text is not the same as the original
                        # replace the text and increment the number of occurrences
                        run.text = replaced_text
                        occurrences[to_replace] += 1
                        
        # Loop through tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # Loop through paragraphs
                    for paragraph in cell.paragraphs:
                        # Loop through runs (style spans)
                        for run in paragraph.runs:
                            # if there is text on this run, replace it
                            if run.text:
                                # get the replacement text
                                replaced_text = re.sub(to_replace, replacement.text, run.text, 999)
                                if replaced_text != run.text:
                                    # if the replaced text is not the same as the original
                                    # replace the text and increment the number of occurrences
                                    run.text = replaced_text
                                    occurrences[to_replace] += 1
                    
        # print the number of occurrences of each word
        for word, count in occurrences.items():
            print(f"The word {word} was found and replaced {count} time(s).")

        # make a new file name by changing the original file name with current word
        index = template_path.rfind('\\') + 1
        new_file_path = template_path[:index] + replacement.text + ".docx"
        print('File was saved at ' + new_file_path)
        # save the new docx file
        doc.save(new_file_path)
else:
    print('The file type is invalid, only .docx are supported')

input('\nPress any button to exit...')
