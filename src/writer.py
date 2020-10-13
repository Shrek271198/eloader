def clear_output():
    import os, re, os.path

    mypath = "output"
    for root, dirs, files in os.walk(mypath):
        for file in files:
            os.remove(os.path.join(root, file))


def mcc_writer(mccl, STANDARD):
    """This method accepts a list of MCC objects and using the MCC
    template produces a summary for each MCC."""

    clear_output()

    from openpyxl import load_workbook
    from shutil import copyfile
    from sys import exit, exc_info

    mcc_template = "src/templates/mcc_template.xlsx"

    for mcc in mccl:

        # Copy and Create MCC excel files
        mcc_output_file = "output/ELECTRICAL LOADS LIST SUBSTATION - {}.xlsx".format(mcc.name)
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
        ws = wb.active

        # Project Name
        ws["G1"] = STANDARD.project_name
        wb.save(filename=mcc_output_file)
