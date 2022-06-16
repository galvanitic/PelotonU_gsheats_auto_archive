from datetime import date
import sched, time
import pygsheets

class HCNSheetsDriver:
    '''
    Main HCN Sheets driver.
    :param main_sheet_id: The id of the main sheet to work on.
    :param archive_folder_id: The id of the folder where all archives will be stored.
    '''
    def __init__(self, main_sheet_id, archive_folder_id):
        # Authorize and create instance
        self.gc = pygsheets.authorize(client_secret='client_secret.json')
        # Global Constants
        self.HCN_ONEGOAL_SHEET_ID = main_sheet_id
        self.ONEGOAL_ARCHIVE_FOLDER_ID = archive_folder_id


    def createBackup(self):
        '''
        Creates a backup of the main sheet.
        :return: Void
        '''
        email_recipients = ['rgalvan@pelotonu.org']
        time_of_backup = date.today().strftime("%m/%d/%Y")
        new_sheet_name = "Archive created: "+time_of_backup
        self.gc.drive.copy_file(self.HCN_ONEGOAL_SHEET_ID, new_sheet_name, self.ONEGOAL_ARCHIVE_FOLDER_ID);
        print(' >> New backup created at '+time_of_backup)
        print(' >> Sharing the archive with folks. This might take a while...')
        sheet_to_share = self.gc.open(new_sheet_name)
        for user in email_recipients:
            sheet_to_share.share(user, emailMessage='PelotonU Bot: This is an auto-generated message to inform you that a new "OneGoal" archive was created. Check it out here!')
            print("     >Sheet was shared with "+ user)
        print("\n \n")

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
    mainSheetID = '1CaaeH03NZOxjQVwQZDusU89Uw_-nqi3RkLJFElziwnk'
    archiveFolderID = '1ocgNSorHOp045yh1XbqIzZ2ALmf65Fvd'
    print("Program is starting...")
    two_weeks = 1209600 # 2 weeks in seconds
    # Create new instance of driver.
    hcn = HCNSheetsDriver(mainSheetID, archiveFolderID)

    def batch():
        print("> Creating backup of main sheet...")
        hcn.createBackup()
        print("> Clearing the main sheet.")
        # hcn.clearTemplate('1ER4WlZ9vMVOYW904DgCvThjr-3IedW4AGHJ9_cwIkvY')

    # define the countdown func.
    def countdownHelper(seconds, granularity=2):
        intervals = (
            ('weeks', 604800),  # 60 * 60 * 24 * 7
            ('days', 86400),    # 60 * 60 * 24
            ('hours', 3600),    # 60 * 60
            ('minutes', 60),
            ('seconds', 1),
        )
        result = []

        for name, count in intervals:
            value = seconds // count
            if value:
                seconds -= value * count
                if value == 1:
                    name = name.rstrip('s')
                result.append("{} {}".format(value, name))
        return ', '.join(result[:granularity])

    def countdown(t):
        print("Next archive will be made in: ")
        while t:
            print(countdownHelper(t, 5), end="\r")
            time.sleep(1)
            t -= 1

    while (True): # Make this run indefinitely
        countdown(5)
        batch()

if __name__ == '__main__':
    main()