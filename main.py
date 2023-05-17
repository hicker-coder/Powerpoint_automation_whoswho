import pandas as pd
import pptx
from datetime import datetime
from presentation_generator import generate_presentation_from_template


if __name__ =='__main__' :
    # open the PowerPoint presentation
    prs = pptx.Presentation("./input_data/template_whoswho_2023.pptx")
    test_excel_path = r'./input_data/raw_data_Mapping_added.xlsx'
    # read the employee data from the Excel file
    raw_df = pd.read_excel(test_excel_path)
    prs = generate_presentation_from_template(prs,raw_df)

    # Get today's date
    today = datetime.now().strftime('%m%d%Y')

    # Update the file path with today's date
    file_path = f'./output_data/auto_generated_whoswho_{today}.pptx'
    prs.save(file_path)






