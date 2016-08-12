Module modRepairNotes

    Private objForm, oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oMatrix As SAPbouiCOM.Matrix
    Private oGrid As SAPbouiCOM.Grid
    Private oCheck As SAPbouiCOM.CheckBox
    Private sSQL As String
    Private oRecordSet As SAPbobsCOM.Recordset
    Private iParentLine As Integer
    Private sParentForm As String
    Private sItemString As String

#Region "Form Initialization from quote"
    Public Sub ReprNotesInitializationFromQuote(ByVal sCardCode As String, ByVal sItemCode As String, ByVal sQuoteDocNo As String, ByVal sQuoteSeries As String, ByVal sLine As String, ByVal sDocType As String, ByVal sToolCatNo As String, ByVal sFormMode As String)
        Dim sFuncName As String = "ReprNotesInitializationFromQuote"
        Dim sErrDesc As String = String.Empty
        Try
            LoadFromXML("Repair Notes.srf", p_oSBOApplication)
            objForm = p_oSBOApplication.Forms.Item("REPR")
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

            objForm.DataSources.DataTables.Add("dtReprNotes")

            GenerateDocNum(objForm)

            LoadDefValues(objForm, sCardCode, sItemCode, sQuoteDocNo, sQuoteSeries, sLine, sDocType, sToolCatNo, sFormMode)

            LoadGrid(objForm, "", sToolCatNo)

            'AddChooseFromList(objForm)
            'AddItemCFLCondition(objForm)
            'DataBinding(objForm)

            'oMatrix = objForm.Items.Item("17").Specific
            'oMatrix.AddRow(1)
            'oMatrix.AutoResizeColumns()

            'objForm.PaneLevel = 3

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
    Public Sub ReprNotes_OpenFormFindMode(ByVal sDocNum As String, ByVal sToolCatDocNo As String, ByVal sLine As String, ByVal sForm As String)
        Dim sFuncName As String = "ReprNotes_OpenFormFindMode"
        Dim sErrDesc As String = String.Empty
        Try
            LoadFromXML("Repair Notes.srf", p_oSBOApplication)
            objForm = p_oSBOApplication.Forms.Item("REPR")
            objForm.Visible = True
            'objForm.Freeze(True)
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

            'oMatrix = objForm.Items.Item("17").Specific
            'oMatrix.AutoResizeColumns()

            'AddChooseFromList(objForm)

            objForm.DataSources.DataTables.Add("dtReprNotes")

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            objForm.PaneLevel = 1
            oEdit = objForm.Items.Item("16").Specific
            oEdit.Value = sDocNum
            iParentLine = sLine
            sParentForm = sForm

            objForm.Items.Item("18").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.PaneLevel = "3"
            LoadGrid(objForm, sDocNum, sToolCatDocNo)
            GetItemString(objForm)

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                If sParentForm = "OQUT" Then
                    oForm = p_oSBOApplication.Forms.GetForm("149", 1)
                    oMatrix = oForm.Items.Item("38").Specific
                    oMatrix.Columns.Item("U_AB_RNote").Cells.Item(iParentLine).Specific.value = sItemString
                ElseIf sParentForm = "OINV" Then
                    oForm = p_oSBOApplication.Forms.GetForm("133", 1)
                    oMatrix = oForm.Items.Item("38").Specific
                    oMatrix.Columns.Item("U_AB_RNote").Cells.Item(iParentLine).Specific.value = sItemString
                End If
            End If

            'objForm.PaneLevel = 2
            'AddItemCFLCondition(objForm)
            'DataBinding(objForm)
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

        sSQL = "SELECT ISNULL(MAX(U_DOCNUM),0) + 1 [DOCNUM] FROM [@AE_REPR] "
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
    Private Sub LoadDefValues(ByVal objForm As SAPbouiCOM.Form, ByVal sCardCode As String, ByVal sItemCode As String, ByVal sQuoteDocNo As String, ByVal sQuoteSeries As String, ByVal sline As String, ByVal sDocType As String, ByVal sToolCatNo As String, ByVal sFormMode As String)
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
        oEdit = objForm.Items.Item("22").Specific
        oEdit.Value = sToolCatNo
        objForm.Items.Item("18").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        objForm.PaneLevel = 3

        objForm.Freeze(False)
    End Sub
#End Region
#Region "Check all Fields"
    Private Function CheckAllFields(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Boolean
        Dim sFuncName As String = "CheckAllFields"
        Dim bCheck As Boolean
        bCheck = True
        sErrDesc = ""

        oGrid = objForm.Items.Item("23").Specific
        Dim v_Count As Integer
        Dim objCheckboxCol As SAPbouiCOM.CheckBoxColumn = oGrid.Columns.Item("Select")
        For i = 0 To oGrid.Rows.Count - 1
            If objCheckboxCol.IsChecked(i) = True Then
                v_Count = v_Count + 1
                Exit For
            End If
        Next
        If v_Count = 0 Then
            bCheck = False
            sErrDesc = "Atleast One Row Selected in Matrix"
            Return bCheck
            Exit Function
        End If

        Return bCheck
    End Function
#End Region
#Region "Choose From List"
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        
        oCFLs = objForm.ChooseFromLists
        oCFLCreationParams = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
        oCFLCreationParams.MultiSelection = False

        'ITEM CFL
        oCFLCreationParams.ObjectType = "4"
        oCFLCreationParams.UniqueID = "CFL1"
        oCFL = oCFLs.Add(oCFLCreationParams)
       

    End Sub
    Private Sub AddItemCFLCondition(ByVal objForm As SAPbouiCOM.Form)
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        oCFLs = objForm.ChooseFromLists
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        oCFLCreationParams = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
        oCFLCreationParams.MultiSelection = True

        Dim sToolsCatNo As String
        oEdit = objForm.Items.Item("22").Specific
        sToolsCatNo = oEdit.Value

        oCFL = oCFLs.Item("CFL1")
        oCons = New SAPbouiCOM.Conditions()

        Dim i As Integer
        i = 0

        sSQL = "EXEC AE_LOADREPAIRNOTES '" & sToolsCatNo & "'"

        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery(sSQL)
        If Not (oRecordSet.BoF And oRecordSet.EoF) Then
            oRecordSet.MoveFirst()
            Do Until oRecordSet.EoF
                i = i + 1
                If i > 1 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                ElseIf i = 1 Then
                    oCon = oCons.Add()
                    oCon.Alias = "frozenFor"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "N"
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                End If

                oCon = oCons.Add()
                oCon.Alias = "ItemName"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRecordSet.Fields.Item("ItemName").Value
                oRecordSet.MoveNext()
            Loop
        Else
            oCon = oCons.Add()
            oCon.Alias = "ItemName"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = ""
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
        oCon = oCons.Add()
        oCon.Alias = "frozenFor"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "N"
        oCFL.SetConditions(oCons)

    End Sub

    Private Sub DataBinding(ByRef objForm As SAPbouiCOM.Form)
        'To Warehouse
        oMatrix = objForm.Items.Item("17").Specific
        oMatrix.Columns.Item("V_1").ChooseFromListUID = "CFL1"
        oMatrix.Columns.Item("V_1").ChooseFromListAlias = "ItemName"
    End Sub
#End Region
#Region "Clear matrix rows"
    Private Sub ClearMatrixRows(ByVal objForm As SAPbouiCOM.Form)
        oMatrix = objForm.Items.Item("17").Specific
        oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific.value = ""
        oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.value = ""

    End Sub
#End Region

#Region "Load Grid Values"
    Private Sub LoadGrid(ByVal objForm As SAPbouiCOM.Form, ByVal sDocNum As String, ByVal sToolCatNo As String)
       
        sSQL = "EXEC AE_LOADREPAIRNOTES '" & sToolCatNo & "','" & sDocNum & "' "
        oGrid = objForm.Items.Item("23").Specific
        objForm.DataSources.DataTables.Item("dtReprNotes").Rows.Clear()
        objForm.DataSources.DataTables.Item("dtReprNotes").ExecuteQuery(sSQL)
        oGrid.DataTable = objForm.DataSources.DataTables.Item("dtReprNotes")
        oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

    End Sub
#End Region
#Region "Move Item to Matrix"
    Private Sub MoveItemsToMatrix(ByVal objForm As SAPbouiCOM.Form)
        Dim iLine, i As Integer
        sItemString = String.Empty

        objForm.PaneLevel = "2"
        oGrid = objForm.Items.Item("23").Specific
        oMatrix = objForm.Items.Item("17").Specific
        oMatrix.Clear()
        iLine = 1

        For i = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue("Select", i) = "Y" Then
                oMatrix.AddRow(1)
                If sItemString = "" Then
                    sItemString = oGrid.DataTable.GetValue("ItemName", i)
                Else
                    sItemString = sItemString & "," & oGrid.DataTable.GetValue("ItemName", i)
                End If
                oMatrix.Columns.Item("V_-1").Cells.Item(iLine).Specific.value = iLine
                oMatrix.Columns.Item("V_1").Cells.Item(iLine).Specific.value = oGrid.DataTable.GetValue("ItemName", i)
                oMatrix.Columns.Item("V_0").Cells.Item(iLine).Specific.value = oGrid.DataTable.GetValue("ForeignName", i)
                iLine = iLine + 1
            End If
        Next

    End Sub
#End Region
#Region "Get Item String"
    Private Sub GetItemString(ByVal objForm As SAPbouiCOM.Form)
        sItemString = String.Empty

        objForm.PaneLevel = "3"
        oGrid = objForm.Items.Item("23").Specific
        
        For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue("Select", i) = "Y" Then
                If sItemString = "" Then
                    sItemString = oGrid.DataTable.GetValue("ItemName", i)
                Else
                    sItemString = sItemString & "," & oGrid.DataTable.GetValue("ItemName", i)
                End If
            End If
        Next

    End Sub
#End Region

#Region "Item Event"
    Public Sub RepairNotes_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "RepairNotes_SBO_ItemEvent"
        Dim sErrDesc As String = String.Empty
        Try
            If pval.Before_Action = True Then
                Select Case pval.EventType
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
                                    MoveItemsToMatrix(objForm)
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
                                    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iLine).Specific.value = sDocNum
                                    oMatrix.Columns.Item("U_AB_RNote").Cells.Item(iLine).Specific.value = sItemString
                                ElseIf sDocType = "13" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iLine).Specific.value = sDocNum
                                    oMatrix.Columns.Item("U_AB_RNote").Cells.Item(iLine).Specific.value = sItemString
                                End If
                            ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                If CheckAllFields(objForm, sErrDesc) = False Then
                                    p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    MoveItemsToMatrix(objForm)
                                End If

                                Dim iLine As Integer
                                Dim sDocType, sDocNum As String
                                oEdit = objForm.Items.Item("20").Specific
                                iLine = oEdit.Value
                                oEdit = objForm.Items.Item("10").Specific
                                sDocType = oEdit.Value
                                oEdit = objForm.Items.Item("16").Specific
                                sDocNum = oEdit.Value
                                'If sDocType = "23" Then
                                '    oForm = p_oSBOApplication.Forms.GetForm("149", pval.FormTypeCount)
                                '    oMatrix = oForm.Items.Item("38").Specific
                                '    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iLine).Specific.value = sDocNum
                                'ElseIf sDocType = "13" Then
                                '    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                '    oMatrix = oForm.Items.Item("38").Specific
                                '    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iLine).Specific.value = sDocNum
                                'End If
                                If sParentForm = "OQUT" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("149", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iParentLine).Specific.value = sDocNum
                                    oMatrix.Columns.Item("U_AB_RNote").Cells.Item(iLine).Specific.value = sItemString
                                ElseIf sParentForm = "OINV" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iParentLine).Specific.value = sDocNum
                                    oMatrix.Columns.Item("U_AB_RNote").Cells.Item(iLine).Specific.value = sItemString
                                End If
                            End If
                        End If

                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.CharPressed = "9" Then
                            If pval.ItemUID = "17" Then
                                oMatrix = objForm.Items.Item("17").Specific
                                If pval.ColUID = "V_1" Then
                                    If pval.Row = oMatrix.RowCount Then
                                        If oMatrix.Columns.Item("V_1").Cells.Item(pval.Row).Specific.value <> "" Then
                                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                objForm.Freeze(True)
                                                oMatrix.AddRow(1)
                                                ClearMatrixRows(objForm)
                                                objForm.Freeze(False)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

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
                                    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iLine).Specific.value = sDocNum
                                ElseIf sDocType = "13" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iLine).Specific.value = sDocNum
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
                                'If sDocType = "23" Then
                                '    oForm = p_oSBOApplication.Forms.GetForm("149", pval.FormTypeCount)
                                '    oMatrix = oForm.Items.Item("38").Specific
                                '    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iLine).Specific.value = sDocNum
                                'ElseIf sDocType = "13" Then
                                '    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                '    oMatrix = oForm.Items.Item("38").Specific
                                '    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iLine).Specific.value = sDocNum
                                'End If
                                If sParentForm = "OQUT" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("149", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iParentLine).Specific.value = sDocNum
                                ElseIf sParentForm = "OINV" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iParentLine).Specific.value = sDocNum
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
                               
                                If sParentForm = "OQUT" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("149", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iParentLine).Specific.value = sDocNum
                                    oMatrix.Columns.Item("U_AB_RNote").Cells.Item(iParentLine).Specific.value = sItemString
                                ElseIf sParentForm = "OINV" Then
                                    oForm = p_oSBOApplication.Forms.GetForm("133", pval.FormTypeCount)
                                    oMatrix = oForm.Items.Item("38").Specific
                                    oMatrix.Columns.Item("U_REPAIRNOTES").Cells.Item(iParentLine).Specific.value = sDocNum
                                    oMatrix.Columns.Item("U_AB_RNote").Cells.Item(iParentLine).Specific.value = sItemString
                                End If
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pval
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        'Dim objForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = objForm.ChooseFromLists.Item(sCFL_ID)
                        Try
                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects
                                If pval.ItemUID = "17" Then
                                    oMatrix = objForm.Items.Item("17").Specific
                                    If pval.Row = 1 Then
                                        oMatrix.Columns.Item("V_-1").Cells.Item(pval.Row).Specific.value = 1
                                    Else
                                        Dim iSno As Integer
                                        iSno = oMatrix.Columns.Item("V_-1").Cells.Item(pval.Row - 1).Specific.value
                                        iSno = iSno + 1
                                        oMatrix.Columns.Item("V_-1").Cells.Item(pval.Row).Specific.value = iSno
                                    End If
                                    If pval.ColUID = "V_1" Then
                                        oMatrix.Columns.Item("V_0").Cells.Item(pval.Row).Specific.value = oDataTable.GetValue("FrgnName", 0)
                                        oMatrix.Columns.Item("V_1").Cells.Item(pval.Row).Specific.value = oDataTable.GetValue("ItemName", 0)
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objForm.Update()
                        End Try

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
#Region "Menu Events"
    Public Sub RepairNotes_SBO_Appln_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                If pVal.MenuUID = "1292" Then
                    objForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.UniqueID)
                    objForm.EnableMenu("1282", False)
                    oMatrix = objForm.Items.Item("17").Specific
                    If oMatrix.RowCount = 0 Then
                        objForm.Freeze(True)
                        oMatrix.AddRow(1)
                        ClearMatrixRows(objForm)
                        objForm.Freeze(False)
                    ElseIf oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.value <> "" Then
                        If pVal.row = oMatrix.RowCount Then
                            objForm.Freeze(True)
                            oMatrix.AddRow(1)
                            ClearMatrixRows(objForm)
                            objForm.Freeze(False)
                        End If
                    End If

                End If
                objForm.Update()
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objForm.Update()
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
            GC.Collect()
        End Try
    End Sub
#End Region

End Module
