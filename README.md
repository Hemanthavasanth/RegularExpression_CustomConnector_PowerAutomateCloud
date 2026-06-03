SET str_ConfigFilePath TO $'''%' '%'''
SET str_ConfigFilePath TO in_ConfigFilePath
BLOCK 'Read excel config file and load in object variable'
ON BLOCK ERROR
    GOTO ExitSubflow
END
    System.TerminateProcess.TerminateProcessByNameWithUserName ProcessUser: $'''recodesolutions\\ers1217''' ProcessName: $'''EXCEL'''
    SET objConfig TO {{ }}
    Excel.LaunchExcel.LaunchAndOpenUnderExistingProcess Path: str_ConfigFilePath Visible: False ReadOnly: False UseMachineLocale: False Instance=> ExcelInstance
    Variables.CreateNewDatatable InputTable: { } DataTable=> dt_Config
    Excel.GetAllWorksheets Instance: ExcelInstance Worksheets=> SheetNames
    LOOP FOREACH Sheet IN SheetNames
        Excel.SetActiveWorksheet.ActivateWorksheetByName Instance: ExcelInstance Name: Sheet
        Excel.GetFirstFreeRowOnColumn Instance: ExcelInstance Column: 1 FirstFreeRowOnColumn=> FirstFreeRowOnColumn
        Display.ShowMessageDialog.ShowMessage Title: $'''Get Free row''' Message: FirstFreeRowOnColumn Icon: Display.Icon.None Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: True
        Excel.ReadFromExcel.ReadCells Instance: ExcelInstance StartColumn: $'''A''' StartRow: 1 EndColumn: $'''B''' EndRow: FirstFreeRowOnColumn - 1 GetCellContentsMode: Excel.GetCellContentsMode.PlainText FirstLineIsHeader: True RangeValue=> dt_Config
        Variables.DeleteEmptyRowsFromDataTable DataTable: dt_Config
        LOOP FOREACH CurrentRow IN dt_Config
            SET strKey TO CurrentRow['Name']
            IF Sheet = $'''Assets''' THEN
                SET strValue TO CurrentRow['Asset']
            ELSE
                SET strValue TO CurrentRow['Value']
            END
            SET objConfig[strKey] TO strValue
        END
    END
    SET objConfig['EntraID'] TO $'''%'hemantha.vasanth@recodesolutions.com'%'''
    Excel.CloseExcel.Close Instance: ExcelInstance
    SET out_ObjDict TO objConfig
END
LABEL ExitSubflow
EXIT FUNCTION
