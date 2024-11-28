' queries for ISAH database
Function check_ProdBillOfMat(bom_table_name As String, prodheader_dossier_code_list As String) As String
    Dim sql0 As String
    sql0 = "SELECT CAST(ProdHeaderDossierCode AS varchar) AS ProdHeaderDossierCode, MIN(RequiredDate) as min_bom_required_date, " & _
    "MAX(RequiredDate) As max_bom_required_date FROM @1 WHERE ProdHeaderDossierCode in @2" & _
    "GROUP BY ProdHeaderDossierCode " & _
    "ORDER BY ProdHeaderDossierCode"
    sql0 = str.subInStr(sql0, bom_table_name, prodheader_dossier_code_list)
    check_ProdBillOfMat = sql0
End Function

Function join_ISAH_EXPORT_CHECK_PROD_BOM() As String
    Dim sql0 As String, table_name1 As String, table_name2 As String
    table_name1 = main.ISAH_STAGING_SHEET_NAME
    table_name2 = main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET
    sql0 = "SELECT a.ProdHeaderOrdNr, a.ProdHeaderDossierCode, a.next_StartDate_header, b.min_bom_required_date, b.max_bom_required_date " & _
    ", SWITCH(a.[StartDate_header] = b.[max_bom_required_date],1,1=1,0) as check_bom_required_date " & _
    "FROM [@1$] a LEFT JOIN [@2$] b ON a.ProdHeaderDossierCode=b.ProdHeaderDossierCode"
    sql0 = str.subInStr(sql0, table_name1, table_name2)
    join_ISAH_EXPORT_CHECK_PROD_BOM = sql0
End Function









