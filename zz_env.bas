'environment constants like in Py

Global Const HOMEPATH = "G:\My Drive"
Global Const WORKPATH = HOMEPATH & "\work"
Global Const EXCEL_TEMPLATE_PATH = HOMEPATH & "\Programming\excel_templates"
Global Const GITHUBPATH = "C:\Users\jskro\Documents\GitHub"

Global Const VDMI_DATAPATH = WORKPATH & "\VDMI\data"
Global Const VDMI_TESTDATAPATH = WORKPATH & "\VDMI\testdata"
Global Const VDMI_CODEPATH = WORKPATH & "\VDMI\vba"
Global Const VDMI_GITHUB = GITHUBPATH & "\VDMI"

Global Const MSSQL_HOME_CONN_STR = "Driver={ODBC Driver 17 for SQL Server};Server=LAPTOP_JKR\SQLEXPRESS;Database=master;Trusted_Connection=yes;"

' code modules
Global Const MODULES_TO_EXPORT = "a;chrt;clls;ctr;db;dict;dt;fs;m;os;r;str;u;vb;w;zz_env"
Global Const VDMI_MODULES_TO_EXPORT = "main;main_isah_queries;database_control;state_control"
Global Const MODULES_TO_IMPORT = "a.bas;chrt.bas;clls.bas;ctr.bas;db.bas;dict.bas;dt.bas;m.bas;os.bas;r.bas;str.bas;u.bas;vb.bas;w.bas"

Sub export_vb_codemodule_code()
    fs.exportModuleCodes MODULES_TO_EXPORT, zz_env.VDMI_CODEPATH, "txt"
    fs.exportModuleCodes MODULES_TO_EXPORT, zz_env.VDMI_CODEPATH
    fs.exportModuleCodes MODULES_TO_EXPORT, zz_env.VDMI_GITHUB
    fs.exportModuleCodes VDMI_MODULES_TO_EXPORT, zz_env.VDMI_GITHUB
End Sub

Sub createExcelMacroTemplate()
    Dim wbname As String, wb1 As Workbook
    timestamp = dt.format_datetime(Now(), "yyyymmdd")
    wbname = "template_" & timestamp
    w.createMacroEnabledTemplate wbname, zz_env.EXCEL_TEMPLATE_PATH, False
    vb.copyModuleCodes ThisWorkbook, Workbooks(wbname), MODULES_TO_EXPORT
    Set wb1 = Workbooks(wbname)
    wb1.Close True
End Sub
