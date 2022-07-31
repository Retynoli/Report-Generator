# Import re for regex functions
import re
# Import docx to work with .docx files.
from docx import Document

# First, we create an empty lists to hold the path to all of lists files and words
list_paths = []
to_replace = []
result_path = []

# Store file paths from input
number_of_lists = input('Write the lists amount: ')
template_path = input('Write the full path to the template file: ')

for i in range(0, int(number_of_lists)):
   list_paths.append(input(f'\n[{i + 1}] Write the full path to the list file: '))
   to_replace.append(input('Write a word to replace for this list: '))

counter = 0
if template_path.endswith('.docx') and list_paths[0].endswith('.docx'):
    words = Document(list_paths[0])
    # Loop through replacer arguments
    for replacement in words.paragraphs:
        doc = Document(template_path)

        for i in range(0, int(number_of_lists)):
            words = Document(list_paths[i])
            cnt = 0
            # Loop through replacer arguments
            for replacement in words.paragraphs:
                if cnt != counter:
                    cnt += 1
                    continue

                cnt = 0
                print(f"\n[{counter + 1}] Сurrent replacement text: {replacement.text}")

                # initialize the number of occurrences of this word to 0
                occurrences = {to_replace[i]: 0}

                # Loop through paragraphs
                for paragraph in doc.paragraphs:
                    # Loop through runs (style spans)
                    for run in paragraph.runs:
                        # if there is text on this run, replace it
                        if run.text:
                            # get the replacement text
                            replaced_text = re.sub(to_replace[i], replacement.text, run.text, 999)
                            if replaced_text != run.text:
                                # if the replaced text is not the same as the original
                                # replace the text and increment the number of occurrences
                                run.text = replaced_text
                                occurrences[to_replace[i]] += 1

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
                                        replaced_text = re.sub(to_replace[i], replacement.text, run.text, 999)
                                        if replaced_text != run.text:
                                            # if the replaced text is not the same as the original
                                            # replace the text and increment the number of occurrences
                                            run.text = replaced_text
                                            occurrences[to_replace[i]] += 1

                # print the number of occurrences of each word
                for word, count in occurrences.items():
                    print(f"The word {word} was found and replaced {count} time(s).")

                break

        # make a new file name by changing the original file name with current word
        index = template_path.rfind('\\') + 1
        # replace all system reserved symbols if found
        filename = replacement.text.translate({ord(c): "-" for c in '\/:*?"<>|'})
        new_file_path = template_path[:index] + filename + ".docx"
        print(f'File was saved at {new_file_path}')
        result_path.append(new_file_path)
        # save the new docx file
        doc.save(new_file_path)
        counter += 1
else:
    print('The file type is invalid, only .docx are supported')

input('\nPress any button to exit...')
