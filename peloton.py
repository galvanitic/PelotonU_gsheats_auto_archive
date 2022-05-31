from datetime import date
import pygsheets

# Global Constants
HCN_ONEGOAL_SHEET_ID = '1CaaeH03NZOxjQVwQZDusU89Uw_-nqi3RkLJFElziwnk'
HCN_ONEGOAL_SHEET_ID_TEST = '11iHlbJxX_UxAZDxhmHl-0pqnSKq26xbjIS4lC3gW02I'
ONEGOAL_ARCHIVE_FOLDER_ID = '1ocgNSorHOp045yh1XbqIzZ2ALmf65Fvd'



# Authorize and create instance
gc = pygsheets.authorize(local=True)

def createBackup():
    time_of_backup = date.today().strftime("%m/%d/%Y")
    new_sheet_name = "Archive created: "+time_of_backup
    gc.drive.copy_file(HCN_ONEGOAL_SHEET_ID, new_sheet_name, ONEGOAL_ARCHIVE_FOLDER_ID);
    print('New backup created at '+time_of_backup)

def clearTemplate(sheet_id):
    '''
    Clears all unnecessary data from sheet.
    :param sheet_id: The ID of the sheet to clear.
    :return: Void
    '''
    sheet_to_clear = gc.open_by_key(sheet_id)

    # Clear 'Dashboard'
    Dashboard = sheet_to_clear.worksheet(property='index', value=0)
    Dashboard.clear(start='C3', end='H10')
    Dashboard.clear(start='C14', end='H18')

    # Clear 'Detailed Academic Data'
    Detailed_Academic_Data = sheet_to_clear.worksheet(property='index', value=1)
    Detailed_Academic_Data.clear(start='A2', end=None)

    # Clear 'Detailed Recruitment Data'
    Detailed_Recruitment_Data = sheet_to_clear.worksheet(property='index', value=2)
    Detailed_Recruitment_Data.clear(start='A2', end=None)

def main():
    # createBackup()
    clearTemplate(HCN_ONEGOAL_SHEET_ID_TEST)

if __name__ == '__main__':
    main()