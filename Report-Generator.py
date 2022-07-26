# Import re for regex functions
import re
# Import docx to work with .docx files.
from docx import Document

# Store file path from CL Arguments.
pattern_path = input('Write pattern file path: ')
list_path = input('Write list file path: ')
to_replace = input('Write word to replace: ')

if pattern_path.endswith('.docx') and list_path.endswith('.docx'):
    words_to_replace = Document(list_path)
    occurrences = {}
    counter = 0
    
    # Loop through replacer arguments
    for replaceArg in words_to_replace.paragraphs:
        doc = Document(pattern_path)
        print("Current text to replace: " + replaceArg.text)      
        # initialize the number of occurences of this word to 0
        occurrences[to_replace] = 0
        
        # Loop through paragraphs
        for para in doc.paragraphs:
            # Loop through runs (style spans)
            for run in para.runs:
                # if there is text on this run, replace it
                if run.text:
                    # get the replacement text
                    replaced_text = re.sub(to_replace, replaceArg.text, run.text, 999)
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
                    for para in cell.paragraphs:
                        # Loop through runs (style spans)
                        for run in para.runs:
                            # if there is text on this run, replace it
                            if run.text:
                                # get the replacement text
                                replaced_text = re.sub(to_replace, replaceArg.text, run.text, 999)
                                if replaced_text != run.text:
                                    # if the replaced text is not the same as the original
                                    # replace the text and increment the number of occurences
                                    run.text = replaced_text
                                    occurrences[to_replace] += 1
                    
        # print the number of occurences of each word
        for word, count in occurrences.items():
            print(f"The word {word} was found and replaced {count} times.")

        # make a new file name by adding "_new" to the original file name
        new_file_path = pattern_path.replace(".docx", "_new" + str(counter) + ".docx")
        # save the new docx file
        doc.save(new_file_path)

        counter += 1
else:
    print('The file type is invalid, only .docx are supported')

input('Press any button to exit...')