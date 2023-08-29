import comtypes.client

def extract_text_from_powerpoint(input_path, output_path):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(input_path)
    
    text = []

    for slide in presentation.Slides:
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text.append(shape.TextFrame.TextRange.Text)
    
    presentation.Close()
    powerpoint.Quit()

    with open(output_path, "w", encoding="utf-8") as txt_file:
        txt_file.write("\n".join(text))

if __name__ == "__main__":
    input_powerpoint = "C:\\Users\\Dell\\OneDrive\\Desktop\\pptx_to_txt\\weeek-end  27 aout 2023.pptx"  # Replace with your input PowerPoint file
    output_text = "output.txt"       # Replace with the desired output text file

    extract_text_from_powerpoint(input_powerpoint, output_text)
    print("Text extracted and saved to", output_text)
