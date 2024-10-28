import os
import excel2img
from docx.shared import Cm
from docxtpl import InlineImage


def get_baselline_image(context, doc, folderpath, file):
    file_Baseline = file
    if os.path.exists(os.path.join(folderpath, file_Baseline)):
        # Generate Image for BaseLine
        file_Baseline = os.path.join(folderpath, file_Baseline)

        excel2img.export_img(file_Baseline, 'KPIImages/Tables/' + 'baseline_table.png', "Sheet1", "T1:Z12")
        baseline_table = InlineImage(doc, 'KPIImages/Tables/baseline_table.png', width=Cm(15), height=Cm(6))

        context.update({
            'baseline_table': baseline_table
        })
        print("Baseline Table Saved!!")
    else:
        print('No File WFC Output -- Baseline Exists!')
