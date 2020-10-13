def clear_output():
    import os, re, os.path

    mypath = "output"
    for root, dirs, files in os.walk(mypath):
        for file in files:
            os.remove(os.path.join(root, file))


def mcc_writer(mccl, STANDARD):
    """This method accepts a list of MCC objects and using the MCC
    template produces a summary for each MCC.

    It does this by first making a copy of the MCC Template file.

    Then making a modification of the `Template for MCC' sheet and
    appending the appropriate detail.

    Finally, it removes the template sheet and saves the file.
    """

    clear_output()

    from openpyxl import load_workbook
    from shutil import copyfile
    from sys import exit, exc_info
    from copy import copy
    from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties

    mcc_template = "src/templates/mcc_template.xlsx"

    for mcc in mccl:

        # Copy and Create MCC excel files
        mcc_title = "ELECTRICAL LOADS LIST SUBSTATION - {}".format(mcc.name)
        mcc_output_file = "output/{}.xlsx".format(mcc_title)
        try:
            copyfile(mcc_template, mcc_output_file)
        except IOError as e:
            print("Unable to copy file. %s" % e)
            exit(1)
        except:
            print("Unexpected error:", exc_info())
            exit(1)

        # Append MCC data into copied file
        wb = load_workbook(filename=mcc_output_file)
        ws = wb["Template for MCC"]

        # Project Title
        ws["G1"] = STANDARD.project_name
        ws["G4"] = mcc_title

        # Project Details
        ws["R1"] = STANDARD.project
        ws["R2"] = STANDARD.revision
        ws["R3"] = STANDARD.prepared_by
        ws["R4"] = STANDARD.date_prepared
        ws["R5"] = STANDARD.approved_by
        ws["R6"] = STANDARD.date_approved

        # Row style
        count =  1
        for me in mcc.mel:
            # Insert row
            ws.insert_rows(9)
            style_cell = "B" + str(9+count)
            # Apply style
            for row in ws.iter_cols(min_row=9, max_row=9, min_col=1, max_col=19):
                for cell in row:
                    cell.style = copy(ws[style_cell].style)
                    cell.font = copy(ws[style_cell].font)
                    cell.border = copy(ws[style_cell].border)
                    cell.fill = copy(ws[style_cell].fill)
                    cell.number_format = copy(ws[style_cell].number_format)
                    cell.protection = copy(ws[style_cell].protection)
                    cell.alignment = copy(ws[style_cell].alignment)

            # Mechanical Equipments

        ws.print_area = "A1:S{}".format(21+len(mcc.mel))
        ws.title = mcc.name
        wb.save(filename=mcc_output_file)
