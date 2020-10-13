def mcc_writer(mccl):
    """This method accepts a list of MCC objects and using the MCC
    template produces a summary for each MCC."""

    from openpyxl import load_workbook
    from shutil import copyfile
    from sys import exit

    mcc_template = "src/templates/mcc_template.xlsx"

    # Copy and Create MCC excel files
    for mcc in mccl:

        try:
            copyfile(mcc_template, 'output/ELECTRICAL LOADS LIST SUBSTATION - {}.xlsx'.format(mcc.name))
        except IOError as e:
            print("Unable to copy file. %s" % e)
            exit(1)
        except:
            print("Unexpected error:", sys.exc_info())
            exit(1)

        print("\nFile copy done!\n")

#    wb = load_workbook(filename=mel)
#    ws = wb.active

