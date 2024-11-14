'environment constants like in Py
Global Const HOMEPATH = "G:\My Drive;F:\My Drive"
Global Const WORKPATH = "G:\My Drive\work;F:\My Drive\work"
Global Const GITHUBPATH = "C:\Users\jskro\Documents\GitHub;C:\Users\JoelKroodsma\Documents\GitHub"
Global Const MSSQL_HOME_CONN_STR = "Driver={ODBC Driver 17 for SQL Server};Server=LAPTOP_JKR\SQLEXPRESS;Database=master;Trusted_Connection=yes;"

' code modules
Global Const MODULES_TO_EXPORT = "a;chrt;clls;ctr;db;dict;dt;fs;m;os;r;str;u;vb;w;zz_env"
Global Const VDMI_MODULES_TO_EXPORT = "main;main_isah_queries;database_control;state_control;ThisWorkbook;Sheet8;Sheet21;tests"
Global Const MODULES_TO_IMPORT = "a.bas;chrt.bas;clls.bas;ctr.bas;db.bas;dict.bas;dt.bas;m.bas;os.bas;r.bas;str.bas;u.bas;vb.bas;w.bas"
Global Const VDMI_MODULES_TO_IMPORT = "main.bas;main_isah_queries.bas;database_control.bas;state_control.bas;ThisWorkbook.bas;Sheet8.bas;Sheet21.bas;tests.bas"

Sub test_zz_env()
    Debug.Assert zz_env.getVDMIGithub() = "C:\Users\JoelKroodsma\Documents\GitHub\VDMI"
End Sub

' export modules to local Github repo
Sub export_vb_codemodule_code()
    fs.exportModuleCodes MODULES_TO_EXPORT, getVDMIGithub()
    fs.exportModuleCodes VDMI_MODULES_TO_EXPORT, getVDMIGithub()
End Sub

' import-update modules from local Github repo
Sub update_vb_codemodule_code()
    fs.updateCodeModules MODULES_TO_IMPORT, getVDMIGithub()
    fs.updateCodeModules VDMI_MODULES_TO_IMPORT, getVDMIGithub()
End Sub

' import-update modules from passed list
Sub update_vb_specific_modules(code_modules_to_update As String)
    For Each modfile In clls.toCollection(code_modules_to_update, ";")
       fs.updateCodeModules CStr(modfile), getVDMIGithub()
    Next
End Sub

' create excel macro template with latest code modules
Sub createExcelMacroTemplate()
    Dim wbname As String, wb1 As Workbook
    timestamp = dt.format_datetime(Now(), "yyyymmdd")
    wbname = timestamp & "_template"
    w.createMacroEnabledTemplate wbname, getExcelTemplatePath(), False
    vb.copyModuleCodes ThisWorkbook, Workbooks(wbname), MODULES_TO_EXPORT
    Set wb1 = Workbooks(wbname)
    wb1.Close True
End Sub

' helper functions
Function getHomePath() As String
    getHomePath = fs.getFirstValidPath(HOMEPATH)
End Function

Function getGithubPath() As String
    getGithubPath = fs.getFirstValidPath(GITHUBPATH)
End Function

Function getWorkPath() As String
    getWorkPath = fs.getFirstValidPath(WORKPATH)
End Function

' Excel templates, testdata
Function getExcelTemplatePath() As String
    getExcelTemplatePath = os.pathJoin(getHomePath(), "Programming\excel_templates")
End Function

Function getExcelTestDataFolder() As String
    getExcelTestDataFolder = os.pathJoin(getHomePath(), "Programming\excel VBA\testdata")
End Function

Function getExcelTestDataFile() As String
    getExcelTestDataFile = os.pathJoin(getExcelTestDataFolder(), "ISAH_mock_tables.xlsx")
End Function

' VDMI
Function getVDMIGithub() As String
    getVDMIGithub = os.pathJoin(getGithubPath(), "VDMI")
End Function

Function getVDMICodePath() As String
    getVDMICodePath = os.pathJoin(getWorkPath(), "VDMI\vba")
End Function

Function getVDMIDataPath() As String
    getVDMIDataPath = os.pathJoin(getWorkPath(), "VDMI\data")
End Function

Function getVDMITestPath() As String
    getVDMITestPath = os.pathJoin(getWorkPath(), "VDMI\testdata")
End Function




