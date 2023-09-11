import comtypes.client
import os
import shutil

# ---------------------------------------|| Version 2.0 ||--------------------------------------------------
# Note: No error, Suceeded, but the txt files are in the same folder
# ---------------------------------------|| Version 2.0 ||--------------------------------------------------


# Define the folder path containing the PowerPoint files

# folder_path = 'C:\\Users\\HP\\Desktop\\samTech\\Projects\\Back-End\\Python\\PowerPoint_to_txt_script\\Powerpoint_to_txt\\All_pp_files'
folder_path = 'C:\\Users\\HP\\Desktop\\pptx_to_txt\\Input_pptx'

output_path = 'C:\\Users\\HP\\Desktop\\pptx_to_txt\\Output_pptx'

# Initialize the PowerPoint application
powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

# Iterate through files in the folder
for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.endswith('.pptx'):
            pptx_path = os.path.join(root, file)
            txt_path = os.path.splitext(pptx_path)[0] + '.txt'

            # Open the PowerPoint presentation
            presentation = powerpoint.Presentations.Open(pptx_path)

            # Initialize a variable to store text from all slides
            all_text = []

            # Iterate through slides and extract text
            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    if shape.HasTextFrame:
                        text_range = shape.TextFrame.TextRange
                        if text_range.Text:
                            all_text.append(text_range.Text)

            # Save the extracted text to a text file
            with open(txt_path, 'w', encoding='utf-8') as txt_file:
                txt_file.write('\n'.join(all_text))

            # Close the presentation
            presentation.Close()

            print(f'Converted {pptx_path} to {txt_path}')

# Quit PowerPoint application
powerpoint.Quit()


for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.endswith('.txt'):
            txt_path = os.path.join(root, file)
            output_txt_path = os.path.join(output_path, file)
            shutil.move(txt_path, output_txt_path)
            print(f'Moved {txt_path} to {output_txt_path}')