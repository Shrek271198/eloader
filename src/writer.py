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
        ws = wb.active

        # Project Title
        ws["G1"] = STANDARD.project_name
        ws["G4"] = mcc_title

        # Project Details
        ws["T1"] = STANDARD.project
        ws["T2"] = STANDARD.revision
        ws["T3"] = STANDARD.prepared_by
        ws["T4"] = STANDARD.date_prepared
        ws["T5"] = STANDARD.approved_by
        ws["T6"] = STANDARD.date_approved





        wb.save(filename=mcc_output_file)
