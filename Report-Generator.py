# Import re for regex functions
import re
# Import docx to work with .docx files.
from docx import Document

# Store file path from CL Arguments.
template_path = input('Write template file path: ')
list_path = input('Write list file path: ')
to_replace = input('Write word to replace: ')

if template_path.endswith('.docx') and list_path.endswith('.docx'):
    words = Document(list_path)
    occurrences = {}
    
    # Loop through replacer arguments
    for replacement in words.paragraphs:
        doc = Document(template_path)
        print("Сurrent replacement text: " + replacement.text)
        # initialize the number of occurences of this word to 0
        occurrences[to_replace] = 0
        
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
                        # replace the text and increment the number of occurences
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
                                    # replace the text and increment the number of occurences
                                    run.text = replaced_text
                                    occurrences[to_replace] += 1
                    
        # print the number of occurences of each word
        for word, count in occurrences.items():
            print(f"The word {word} was found and replaced {count} times.")

        # make a new file name by changing the original file name with current word
        index = template_path.rfind('\\') + 1
        new_file_path = template_path[:index] + replacement.text + ".docx"
        print('File was saved at ' + new_file_path + '\n')
        # save the new docx file
        doc.save(new_file_path)
else:
    print('The file type is invalid, only .docx are supported')

input('Press any button to exit...')