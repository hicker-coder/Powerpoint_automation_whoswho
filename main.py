import pandas as pd
import pptx
from presentation_generator import generate_presentation_from_template


if __name__ =='__main__' :
    # open the PowerPoint presentation
    prs = pptx.Presentation("./input_data/template_whoswho_2023.pptx")
    test_excel_path = r'./input_data/raw_data_Mapping_added.xlsx'
    # read the employee data from the Excel file
    raw_df = pd.read_excel(test_excel_path)
    prs = generate_presentation_from_template(prs,raw_df)

    prs.save('./output_data/auto_edited_whoswho_04202023.pptx')





