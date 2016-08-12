Module modSalesQuotation

    Private oItem As SAPbouiCOM.Item
    Private oEdit As SAPbouiCOM.EditText
    Private oMatrix As SAPbouiCOM.Matrix
    Private oCheck As SAPbouiCOM.CheckBox
    Private oCombo As SAPbouiCOM.ComboBox
    Private sSql, sDocNo, sSeries As String
    Private oRecordSet As SAPbobsCOM.Recordset

#Region "Open Tools Category Screen"
    Private Sub OpenToolsCategory(ByVal objForm As SAPbouiCOM.Form, ByVal iLine As Integer)
        Dim sQuoteDocNo, sQuoteSeries, sCardCode, sItemCode, sToolCatNo As String
        Dim iCount As Integer

        oEdit = objForm.Items.Item("8").Specific
        sQuoteDocNo = oEdit.Value
        oCombo = objForm.Items.Item("88").Specific
        sQuoteSeries = oCombo.Selected.Value
        oEdit = objForm.Items.Item("4").Specific
        sCardCode = oEdit.Value
        sItemCode = oMatrix.Columns.Item("1").Cells.Item(iLine).Specific.value

        sSql = "SELECT COUNT(*)MNO , U_DOCNUM FROM [@AE_TCSS] WHERE U_CARDCODE = '" & sCardCode & "' AND U_ITEMCODE = '" & sItemCode & "' " & _
               " AND U_BASEDOCNO = '" & sQuoteDocNo & "' AND U_BASEDOCSERIES = '" & sQuoteSeries & "' AND U_ITEMLINE = '" & iLine & "' " & _
               " GROUP BY U_DOCNUM "
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery(sSql)
        If oRecordSet.RecordCount > 0 Then
            iCount = oRecordSet.Fields.Item("MNO").Value
            sToolCatNo = oRecordSet.Fields.Item("U_DOCNUM").Value
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

        If iCount > 0 Then
            ToolCate_OpenFormFindMode(sToolCatNo, objForm.Mode)
        Else
            ToolCatFormInitializationFromQuote(sCardCode, sItemCode, sQuoteDocNo, sQuoteSeries, iLine, "23", objForm.Mode)
        End If

    End Sub
#End Region
#Region "Open Repair Notes Screen"
    Private Sub OpenRepairNotes(ByVal objForm As SAPbouiCOM.Form, ByVal iLine As Integer)
        Dim sQuoteDocNo, sQuoteSeries, sCardCode, sItemCode, sReprNtsDocNo, sToolCatNo As String
        Dim iCount As Integer

        oEdit = objForm.Items.Item("8").Specific
        sQuoteDocNo = oEdit.Value
        oCombo = objForm.Items.Item("88").Specific
        sQuoteSeries = oCombo.Selected.Value
        oEdit = objForm.Items.Item("4").Specific
        sCardCode = oEdit.Value
        sItemCode = oMatrix.Columns.Item("1").Cells.Item(iLine).Specific.value
        sToolCatNo = oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value

        sSql = "SELECT COUNT(*)MNO , U_DOCNUM FROM [@AE_REPR] WHERE U_CARDCODE = '" & sCardCode & "' AND U_ITEMCODE = '" & sItemCode & "' " & _
               " AND U_BASEDOCNO = '" & sQuoteDocNo & "' AND U_BASEDOCSERIES = '" & sQuoteSeries & "' AND U_ITEMLINE = '" & iLine & "' " & _
               " GROUP BY U_DOCNUM "
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery(sSql)
        If oRecordSet.RecordCount > 0 Then
            iCount = oRecordSet.Fields.Item("MNO").Value
            sReprNtsDocNo = oRecordSet.Fields.Item("U_DOCNUM").Value
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

        If iCount > 0 Then
            ReprNotes_OpenFormFindMode(sReprNtsDocNo, sToolCatNo, iLine, "OQUT")
        Else
            ReprNotesInitializationFromQuote(sCardCode, sItemCode, sQuoteDocNo, sQuoteSeries, iLine, "23", sToolCatNo, objForm.Mode)
        End If

    End Sub
#End Region
#Region "Delete Sub Forms Data"
    Private Sub DeleteSubFormData(ByVal objForm As SAPbouiCOM.Form)
        Dim sToolCatNo, sReprNtsDocNo As String
        oMatrix = objForm.Items.Item("38").Specific

        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        For i As Integer = 1 To oMatrix.RowCount
            If oMatrix.Columns.Item("1").Cells.Item(i).Specific.value <> "" Then
                sToolCatNo = oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(i).Specific.value

                If sToolCatNo <> "" Then
                    sSql = "DELETE FROM [@AE_TCS1] WHERE DocEntry = (SELECT DocEntry FROM [@AE_TCSS] WHERE U_DOCNUM = '" & sToolCatNo & "')"
                    oRecordSet.DoQuery(sSql)

                    sSql = "DELETE FROM [@AE_TCSS] WHERE U_DOCNUM = '" & sToolCatNo & "'"
                    oRecordSet.DoQuery(sSql)
                End If

                sReprNtsDocNo = oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(i).Specific.value
                If sReprNtsDocNo <> "" Then
                    sSql = "DELETE FROM [@AE_EPR1] WHERE DocEntry = (SELECT DocEntry FROM [@AE_REPR] WHERE U_DOCNUM = '" & sReprNtsDocNo & "')"
                    oRecordSet.DoQuery(sSql)

                    sSql = "DELETE FROM [@AE_REPR] WHERE U_DOCNUM = '" & sReprNtsDocNo & "'"
                    oRecordSet.DoQuery(sSql)
                End If

            End If
        Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

    End Sub
#End Region
#Region "Delete Unchecked values in Tools Category Selection"
    Private Sub DelUncheckValues(ByVal objForm As SAPbouiCOM.Form)
        sSql = "DELETE FROM [@AE_TCS1]  WHERE U_SELECT = 'N' " & _
               " AND DocEntry = (SELECT DocEntry FROM [@AE_TCSS] WHERE U_BASEDOCNO = '" & sDocNo & "' AND U_BASEDOCSERIES = '" & sSeries & "' AND U_DOCTYPE = '23') "
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery(sSql)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
    End Sub
#End Region

#Region "Item Event"
    Public Sub SalesQuote_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "BP_SBO_ItemEvent"
        Dim sErrDesc As String = String.Empty
        Try
            If pval.Before_Action = True Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "1" Then
                            oEdit = objForm.Items.Item("8").Specific
                            sDocNo = oEdit.Value
                            oCombo = objForm.Items.Item("88").Specific
                            sSeries = oCombo.Selected.Value
                        ElseIf pval.ItemUID = "2" Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                DeleteSubFormData(objForm)
                            End If
                        End If

                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "1" Then
                            If pval.Action_Success = True Then
                                'DelUncheckValues(objForm)
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "38" Then
                            oMatrix = objForm.Items.Item("38").Specific
                            If pval.Row > 0 Then
                                If pval.Row <= oMatrix.RowCount And oMatrix.Columns.Item("1").Cells.Item(pval.Row).Specific.value <> "" Then
                                    If pval.ColUID = "U_TOOLSCATEGORY" Then
                                        OpenToolsCategory(objForm, pval.Row)
                                    ElseIf pval.ColUID = "U_REPAIRNOTES" Then
                                        If oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(pval.Row).Specific.value <> "" Then
                                            OpenRepairNotes(objForm, pval.Row)
                                        Else
                                            p_oSBOApplication.StatusBar.SetText("Select Tools category first", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        End If
                                    End If
                                End If

                            End If
                        End If

                End Select
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region
#Region "Menu Event"
    Public Sub BP_SBO_MenuEvent(ByVal pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form
                If pVal.MenuUID = "AE_VM" Then

                ElseIf pVal.MenuUID = "1282" Then
                    objForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.UniqueID)

                ElseIf pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291" Then
                    objForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.UniqueID)

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
#End Region

End Module
