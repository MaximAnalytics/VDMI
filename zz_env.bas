'environment constants like in Py
Global Const HOMEPATH = "G:\My Drive;F:\My Drive"
Global Const WORKPATH = "G:\My Drive\work;F:\My Drive\work"
Global Const GITHUBPATH = "C:\Users\jskro\Documents\GitHub;C:\Users\JoelKroodsma\Documents\GitHub"
Global Const MSSQL_HOME_CONN_STR = "Driver={ODBC Driver 17 for SQL Server};Server=LAPTOP_JKR\SQLEXPRESS;Database=master;Trusted_Connection=yes;"

' code modules
Global Const MODULES_TO_EXPORT = "a;chrt;clls;ctr;db;dict;dt;fs;m;os;r;str;u;vb;w;zz_env"
Global Const VDMI_MODULES_TO_EXPORT = "main;main_isah_queries;database_control;state_control;ThisWorkbook;Sheet8;tests"
Global Const MODULES_TO_IMPORT = "a.bas;chrt.bas;clls.bas;ctr.bas;db.bas;dict.bas;dt.bas;m.bas;os.bas;r.bas;str.bas;u.bas;vb.bas;w.bas"

Sub test_zz_env()
    Debug.Assert zz_env.getVDMIGithub() = "C:\Users\JoelKroodsma\Documents\GitHub\VDMI"
End Sub

' export modules and export as template
Sub export_vb_codemodule_code()
    'fs.exportModuleCodes MODULES_TO_EXPORT, getVDMICodePath(), "txt"
    'fs.exportModuleCodes MODULES_TO_EXPORT, getVDMICodePath()
    fs.exportModuleCodes MODULES_TO_EXPORT, getVDMIGithub()
    fs.exportModuleCodes VDMI_MODULES_TO_EXPORT, getVDMIGithub()
End Sub

Sub update_vb_codemodule_code()
    fs.updateCodeModules "fs.bas", getVDMIGithub()
End Sub

Sub createExcelMacroTemplate()
    Dim wbname As String, wb1 As Workbook
    timestamp = dt.format_datetime(Now(), "yyyymmdd")
    wbname = "template_" & timestamp
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

Function getVDMIGithub() As String
    getVDMIGithub = os.pathJoin(getGithubPath(), "VDMI")
End Function

Function getWorkPath() As String
    getWorkPath = fs.getFirstValidPath(WORKPATH)
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

Function getExcelTemplatePath() As String
    getExcelTemplatePath = fs.getFirstValidPath(getHomePath(), "Programming\excel_templates")
End Function


