Module modToolCategory

    Private objForm, oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oMatrix As SAPbouiCOM.Matrix
    Private oCheck As SAPbouiCOM.CheckBox
    Private sSQL As String
    Private oRecordSet As SAPbobsCOM.Recordset

#Region "Form Initialization from quote"
    Public Sub ToolCatFormInitializationFromQuote(ByVal sCardCode As String, ByVal sItemCode As String, ByVal sQuoteDocNo As String, ByVal sQuoteSeries As String, ByVal sLine As String, ByVal sDocType As String, ByVal sFormMode As String)
        Dim sFuncName As String = "FormInitializationFromQuote"
        Dim sErrDesc As String = String.Empty
        Try
            LoadFromXML("Tools Category Selection.srf", p_oSBOApplication)
            objForm = p_oSBOApplication.Forms.Item("TCS")
            objForm.Visible = True
            objForm.Freeze(True)
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            objForm.EnableMenu("6913", False) 'User Defined windows
            objForm.EnableMenu("1290", False) 'Move First Record
            objForm.EnableMenu("1288", False) 'Move Next Record
            objForm.EnableMenu("1289", False) 'Move Previous Record
            objForm.EnableMenu("1291", False) 'Move Last Record
            objForm.EnableMenu("1281", True) 'Find Record
            objForm.EnableMenu("1282", False) 'Add New Record
            objForm.EnableMenu("1292", False) 'Add New Row

            objForm.DataBrowser.BrowseBy = "16"

            GenerateDocNum(objForm)

            LoadDefValues(objForm, sCardCode, sItemCode, sQuoteDocNo, sQuoteSeries, sLine, sDocType, sFormMode)
            LoadMatrix(objForm, "", sFormMode)

            oMatrix = objForm.Items.Item("17").Specific
            ' oMatrix.AddRow(1)
            oMatrix.AutoResizeColumns()

            objForm.Freeze(False)
            objForm.Update()
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region
#Region "Open Form in Find Mode"
    Public Sub ToolCate_OpenFormFindMode(ByVal sDocNum As String, ByVal sFormMode As String)
        Dim sFuncName As String = "OpenFormFindMode"
        Dim sErrDesc As String = String.Empty
        Try
            LoadFromXML("Tools Category Selection.srf", p_oSBOApplication)
            objForm = p_oSBOApplication.Forms.Item("TCS")
            objForm.Visible = True
            objForm.Freeze(True)
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            objForm.EnableMenu("6913", False) 'User Defined windows
            objForm.EnableMenu("1290", False) 'Move First Record
            objForm.EnableMenu("1288", False) 'Move Next Record
            objForm.EnableMenu("1289", False) 'Move Previous Record
            objForm.EnableMenu("1291", False) 'Move Last Record
            objForm.EnableMenu("1281", True) 'Find Record
            objForm.EnableMenu("1282", False) 'Add New Record
            objForm.EnableMenu("1292", False) 'Add New Row

            objForm.DataBrowser.BrowseBy = "16"

            oMatrix = objForm.Items.Item("17").Specific
            oMatrix.AutoResizeColumns()

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            objForm.PaneLevel = 1
            oEdit = objForm.Items.Item("16").Specific
            oEdit.Value = sDocNum
            ' objForm.PaneLevel = 2

            objForm.Items.Item("18").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            LoadMatrix(objForm, sDocNum, sFormMode)

            objForm.PaneLevel = 2

            objForm.Freeze(False)
            objForm.Update()
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region
#Region "Generate Document number"
    Private Sub GenerateDocNum(ByVal objForm As SAPbouiCOM.Form)
        Dim sErrDesc As String = String.Empty
        Dim sSQL As String

        sSQL = "SELECT ISNULL(MAX(U_DOCNUM),0) + 1 [DOCNUM] FROM [@AE_TCSS] "
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery(sSQL)
        If oRecordSet.RecordCount > 0 Then
            oEdit = objForm.Items.Item("16").Specific
            oEdit.Value = oRecordSet.Fields.Item("DOCNUM").Value
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

        objForm.Items.Item("18").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        objForm.Items.Item("14").Enabled = False

    End Sub
#End Region
#Region "Load Predefined values"
    Private Sub LoadDefValues(ByVal objForm As SAPbouiCOM.Form, ByVal sCardCode As String, ByVal sItemCode As String, ByVal sQuoteDocNo As String, ByVal sQuoteSeries As String, ByVal sline As String, ByVal sDocType As String, ByVal sFormMode As String)
        objForm.Freeze(True)
        objForm.PaneLevel = 1
        oEdit = objForm.Items.Item("4").Specific
        oEdit.Value = sCardCode
        oEdit = objForm.Items.Item("7").Specific
        oEdit.Value = sItemCode
        oEdit = objForm.Items.Item("12").Specific
        oEdit.Value = sQuoteDocNo
        oEdit = objForm.Items.Item("14").Specific
        oEdit.Value = sQuoteSeries
        oEdit = objForm.Items.Item("20").Specific
        oEdit.Value = sline
        oEdit = objForm.Items.Item("10").Specific
        oEdit.Value = sDocType
        objForm.Items.Item("18").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        
        objForm.Freeze(False)
    End Sub
#End Region
#Region "Load Matrix values"
    Private Sub LoadMatrix(ByVal objForm As SAPbouiCOM.Form, ByVal sDocNum As String, ByVal sFormMode As String)
        objForm.PaneLevel = 2
        oMatrix = objForm.Items.Item("17").Specific
        If sFormMode = "3" Then
            sSQL = "SELECT ItmsTypCod,ItmsGrpNam FROM OITG ORDER BY ItmsTypCod ASC"
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sSQL)
            If Not (oRecordSet.BoF And oRecordSet.EoF) Then
                oRecordSet.MoveFirst()
                Do Until oRecordSet.EoF
                    oMatrix.AddRow(1)
                    oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.value = oRecordSet.Fields.Item("ItmsTypCod").Value
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific.value = oRecordSet.Fields.Item("ItmsGrpNam").Value
                    oRecordSet.MoveNext()
                Loop
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
        Else
            'objForm.Items.Item("17").Enabled = False
            'oMatrix.Clear()
            'sSQL = "DECLARE @TOOLCATDOCNO NVARCHAR(MAX) "
            'sSQL = sSQL & " SET @TOOLCATDOCNO = '" & sDocNum & "' "
            'sSQL = sSQL & " IF ISNULL(@TOOLCATDOCNO,'') = '' "
            'sSQL = sSQL & " BEGIN "
            'sSQL = sSQL & " SELECT ItmsTypCod,ItmsGrpNam FROM OITG ORDER BY ItmsTypCod ASC "
            'sSQL = sSQL & " END "
            'sSQL = sSQL & " ELSE "
            'sSQL = sSQL & " BEGIN "
            'sSQL = sSQL & " SELECT ISNULL((SELECT 'Y' FROM [@AE_TCSS] A INNER JOIN [@AE_TCS1] B ON B.DocEntry = A.DocEntry WHERE A.U_DOCNUM = @TOOLCATDOCNO AND B.U_ITEMPROPERTYCODE = ItmsTypCod),'') [Select], "
            'sSQL = sSQL & " ItmsTypCod,ItmsGrpNam FROM OITG ORDER BY ItmsTypCod ASC "
            'sSQL = sSQL & " END "
            'oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecordSet.DoQuery(sSQL)
            'If Not (oRecordSet.BoF And oRecordSet.EoF) Then
            '    oRecordSet.MoveFirst()
            '    Do Until oRecordSet.EoF
            '        oMatrix.AddRow(1)
            '        oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.value = oMatrix.RowCount
            '        oCheck = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
            '        If oRecordSet.Fields.Item("Select").Value = "Y" Then
            '            oCheck.Checked = True
            '        Else
            '            oCheck.Checked = False
            '        End If
            '        oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.value = oRecordSet.Fields.Item("ItmsTypCod").Value
            '        oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific.value = oRecordSet.Fields.Item("ItmsGrpNam").Value
            '        oRecordSet.MoveNext()
            '    Loop
            'End If
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

            'objForm.Update()
            'objForm.Items.Item("17").Enabled = True
        End If

    End Sub
#End Region
#Region "Check all Fields"
    Private Function CheckAllFields(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Boolean
        Dim sFuncName As String = "CheckAllFields"
        Dim bCheck As Boolean
        bCheck = True
        sErrDesc = ""

        oMatrix = objForm.Items.Item("17").Specific
        For i = 1 To oMatrix.RowCount
            oMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.value = i
        Next

        For i = 1 To oMatrix.RowCount
            oCheck = oMatrix.Columns.Item("V_2").Cells.Item(i).Specific
            If oCheck.Checked = True Then
                bCheck = True
                Exit For
            Else
                bCheck = False
            End If
        Next

        If bCheck = False Then
            sErrDesc = "Select atleast one row in matrix"
            bCheck = False
            Return bCheck
            Exit Function
        End If

        Return bCheck
    End Function
#End Region

#Region "Item Event"
    Public Sub ToolsCate_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "BP_SBO_ItemEvent"
        Dim sErrDesc As String = String.Empty
        Try
            If pval.Before_Action = True Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.CharPressed = "9" Then

                        End If

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "1" Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If CheckAllFields(objForm, sErrDesc) = False Then
                                    p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    GenerateDocNum(objForm)
                                End If

                                Dim iLine As Integer
                                Dim sDocType, sDocNum As String
                                oEdit = objForm.Items.Item("20").Specific
                                iLine = oEdit.Value
                                oEdit = objForm.Items.Item("10").Specific
                                sDocType = oEdit.Value
                                oEdit = objForm.Items.Item("16").Specific
                                sDocNum = oEdit.Value
                                If sDocType = "23" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("149", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value = sDocNum
                                ElseIf sDocType = "13" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value = sDocNum
                                End If
                            ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                If CheckAllFields(objForm, sErrDesc) = False Then
                                    p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                Dim iLine As Integer
                                Dim sDocType, sDocNum As String
                                oEdit = objForm.Items.Item("20").Specific
                                iLine = oEdit.Value
                                oEdit = objForm.Items.Item("10").Specific
                                sDocType = oEdit.Value
                                oEdit = objForm.Items.Item("16").Specific
                                sDocNum = oEdit.Value
                                If sDocType = "23" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("149", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value = sDocNum
                                ElseIf sDocType = "13" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value = sDocNum
                                End If
                            End If
                        End If

                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "1" Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                Dim iLine As Integer
                                Dim sDocType, sDocNum As String
                                oEdit = objForm.Items.Item("20").Specific
                                Try
                                    iLine = oEdit.Value
                                Catch ex As Exception
                                End Try
                                oEdit = objForm.Items.Item("10").Specific
                                sDocType = oEdit.Value
                                oEdit = objForm.Items.Item("16").Specific
                                sDocNum = oEdit.Value
                                If sDocType = "23" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("149", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value = sDocNum
                                ElseIf sDocType = "13" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value = sDocNum
                                End If
                                objForm.Close()
                            ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                Dim iLine As Integer
                                Dim sDocType, sDocNum As String
                                oEdit = objForm.Items.Item("20").Specific
                                iLine = oEdit.Value
                                oEdit = objForm.Items.Item("10").Specific
                                sDocType = oEdit.Value
                                oEdit = objForm.Items.Item("16").Specific
                                sDocNum = oEdit.Value
                                If sDocType = "23" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("149", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value = sDocNum
                                ElseIf sDocType = "13" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value = sDocNum
                                End If
                                objForm.Close()
                            ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim iLine As Integer
                                Dim sDocType, sDocNum As String
                                oEdit = objForm.Items.Item("20").Specific
                                iLine = oEdit.Value
                                oEdit = objForm.Items.Item("10").Specific
                                sDocType = oEdit.Value
                                oEdit = objForm.Items.Item("16").Specific
                                sDocNum = oEdit.Value
                                If sDocType = "23" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("149", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value = sDocNum
                                ElseIf sDocType = "13" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_TOOLSCATEGORY").Cells.Item(iLine).Specific.value = sDocNum
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

End Module
