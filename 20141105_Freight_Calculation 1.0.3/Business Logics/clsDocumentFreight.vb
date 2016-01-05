Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsDocumentFreight
    Inherits clsBase
    Private oMatrix As SAPbouiCOM.Matrix
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private Const strAirFreight As String = "Air Freight"
    Private Const strCustomDuty As String = "Custom Duty"

    Public Sub New()
        MyBase.New()
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.MenuUID
            Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            Case mnu_ADD
        End Select
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_DocumentFreight And modVariables.frmFBaseType = "17" Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_CLICK, SAPbouiCOM.BoEventTypes.et_KEY_DOWN, SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("3").Specific
                                If pVal.ItemUID = "3" Then
                                    If (pVal.ColUID = "U_Percent") Then
                                        Dim strFreight As String = oApplication.Utilities.getFreightName(CType(oMatrix.Columns.Item("1").Cells().Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value)
                                        If strFreight <> strCustomDuty Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    ElseIf (pVal.ColUID = "U_PRate") Then
                                        Dim strFreight As String = oApplication.Utilities.getFreightName(CType(oMatrix.Columns.Item("1").Cells().Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value)
                                        If strFreight <> strAirFreight Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initializeControls(oForm)
                                initialize(oForm)
                                If modVariables.blnAutoFreight Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    Else
                                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If
                                Else
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("3").Specific
                                If pVal.ItemUID = "3" And (pVal.ColUID = "U_PRate" Or pVal.ColUID = "U_Percent") And pVal.Row > 0 Then
                                    calculateDiscount(oForm, pVal.Row)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                                modVariables.frmFBaseType = ""
                                modVariables.dblSOQuanity = 0
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.BeforeAction
                Case True

                Case False
                    Select Case BusinessObjectInfo.EventType
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try
            
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strFreight, strStatus As String

            If CType(oForm.Items.Item("4").Specific, SAPbouiCOM.CheckBox).Checked Then
                CType(oForm.Items.Item("4").Specific, SAPbouiCOM.CheckBox).Checked = False
            End If

            If oForm.Items.Item("3").Enabled Then
                oForm.Freeze(True)
                For index As Integer = 1 To oMatrix.VisualRowCount
                    strFreight = oApplication.Utilities.getFreightName(CType(oMatrix.Columns.Item("1").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value)
                    strStatus = CType(oMatrix.Columns.Item("10000071").Cells().Item(index).Specific, SAPbouiCOM.ComboBox).Selected.Value
                    If strStatus = "O" Then
                        Dim dblPRate As Double = IIf(CType(oMatrix.Columns.Item("U_PRate").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value <= 0, 4.6, CType(oMatrix.Columns.Item("U_PRate").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value)
                        Dim dblPercent As Double = IIf(CType(oMatrix.Columns.Item("U_Percent").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value <= 0, 5, CType(oMatrix.Columns.Item("U_Percent").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value)
                        If strFreight = strAirFreight Then
                            If CType(oMatrix.Columns.Item("U_PRate").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value <= 0 Then
                                CType(oMatrix.Columns.Item("U_PRate").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = dblPRate
                                CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = dblPRate * modVariables.dblSOQuanity
                            Else
                                CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = dblPRate * modVariables.dblSOQuanity
                            End If
                        ElseIf strFreight = strCustomDuty Then
                            Dim dblAirFreight As Double = dblPRate * modVariables.dblSOQuanity
                            If CType(oMatrix.Columns.Item("U_Percent").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value <= 0 Then
                                CType(oMatrix.Columns.Item("U_Percent").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = dblPercent
                                CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = ((modVariables.dblBfDisTotal - modVariables.dblDiscount) + dblAirFreight) * (dblPercent / 100)
                            Else
                                CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = ((modVariables.dblBfDisTotal - modVariables.dblDiscount) + dblAirFreight) * (dblPercent / 100)
                            End If
                        End If
                    End If
                Next
                oForm.Freeze(False)
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub calculateDiscount(ByVal oForm As SAPbouiCOM.Form, ByVal intRow As Integer)
        Try
            oForm.Freeze(True)
            oMatrix = oForm.Items.Item("3").Specific
            Dim strFreight As String
            Dim dblAirFreight, dblRate As Double
            strFreight = oApplication.Utilities.getFreightName(CType(oMatrix.Columns.Item("1").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value)
            If strFreight = strAirFreight Then
                dblRate = CType(oMatrix.Columns.Item("U_PRate").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value
                dblAirFreight = dblRate * modVariables.dblSOQuanity
                CType(oMatrix.Columns.Item("3").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value = dblAirFreight
                For index As Integer = 1 To oMatrix.VisualRowCount
                    strFreight = oApplication.Utilities.getFreightName(CType(oMatrix.Columns.Item("1").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value)
                    If strFreight = strCustomDuty Then
                        Dim dblPercent As Double = CType(oMatrix.Columns.Item("U_Percent").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                        CType(oMatrix.Columns.Item("3").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value = ((modVariables.dblBfDisTotal - modVariables.dblDiscount) + dblAirFreight) * (dblPercent / 100)
                    End If
                Next
            ElseIf strFreight = strCustomDuty Then
                Dim dblPercent As Double = CType(oMatrix.Columns.Item("U_Percent").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value
                For index As Integer = 1 To oMatrix.VisualRowCount
                    strFreight = oApplication.Utilities.getFreightName(CType(oMatrix.Columns.Item("1").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value)
                    dblRate = CType(oMatrix.Columns.Item("U_PRate").Cells().Item(index).Specific, SAPbouiCOM.EditText).Value
                    dblAirFreight = dblRate * modVariables.dblSOQuanity
                    If strFreight = strAirFreight Then
                        CType(oMatrix.Columns.Item("3").Cells().Item(intRow).Specific, SAPbouiCOM.EditText).Value = ((modVariables.dblBfDisTotal - modVariables.dblDiscount) + dblAirFreight) * (dblPercent / 100)
                    End If
                Next
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#End Region

End Class
