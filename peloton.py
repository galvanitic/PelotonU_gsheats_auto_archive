from datetime import date
import pygsheets

class HCNSheetsDriver:
    def __init__(self, main_sheet_id, archive_folder_id):
        # Authorize and create instance
        self.gc = pygsheets.authorize(local=True)
        # Global Constants
        self.HCN_ONEGOAL_SHEET_ID = main_sheet_id
        self.ONEGOAL_ARCHIVE_FOLDER_ID = archive_folder_id


    def createBackup(self):
        time_of_backup = date.today().strftime("%m/%d/%Y")
        new_sheet_name = "Archive created: "+time_of_backup
        self.gc.drive.copy_file(self.HCN_ONEGOAL_SHEET_ID, new_sheet_name, self.ONEGOAL_ARCHIVE_FOLDER_ID);
        print('New backup created at '+time_of_backup)

    def clearTemplate(self, sheet_id):
        '''
        Clears all unnecessary data from sheet.
        :param sheet_id: The ID of the sheet to clear.
        :return: Void
        '''
        sheet_to_clear = self.gc.open_by_key(sheet_id)

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
    hcn = HCNSheetsDriver('1CaaeH03NZOxjQVwQZDusU89Uw_-nqi3RkLJFElziwnk', '1ocgNSorHOp045yh1XbqIzZ2ALmf65Fvd')
    
    # hcn.createBackup()
    hcn.clearTemplate('1ER4WlZ9vMVOYW904DgCvThjr-3IedW4AGHJ9_cwIkvY')

if __name__ == '__main__':
    main()