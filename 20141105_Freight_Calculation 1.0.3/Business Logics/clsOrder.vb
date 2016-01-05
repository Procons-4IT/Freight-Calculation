Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsOrder
    Inherits clsBase
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Private strQuery As String
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText

    Public Sub New()
        MyBase.New()
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
            End Select
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ORDR Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "91" Then
                                    Dim strDocBefDiscount, strDiscount As String
                                    strDocBefDiscount = oForm.Items.Item("22").Specific.value
                                    strDiscount = oForm.Items.Item("42").Specific.value
                                    CType(oForm.Items.Item("_52").Specific, SAPbouiCOM.EditText).Value = getDocumentTotal(oForm).ToString()
                                    modVariables.dblSOQuanity = getDocumentTotal(oForm)
                                    modVariables.dblBfDisTotal = oApplication.Utilities.getDocumentQuantity(strDocBefDiscount)
                                    modVariables.dblDiscount = oApplication.Utilities.getDocumentQuantity(strDiscount)
                                    modVariables.frmFBaseType = "17"
                                ElseIf pVal.ItemUID = "1" Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If CType(oForm.Items.Item("81").Specific, SAPbouiCOM.ComboBox).Selected.Value = "1" Then
                                            CType(oForm.Items.Item("_52").Specific, SAPbouiCOM.EditText).Value = getDocumentTotal(oForm).ToString()
                                            If Not modVariables.blnAutoFreight Then
                                                ' modVariables.blnAutoFreight = True
                                                ' oForm.Items.Item("91").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                CType(oForm.Items.Item("_52").Specific, SAPbouiCOM.EditText).Value = getDocumentTotal(oForm).ToString()
                                                ' modVariables.blnAutoFreight = False
                                            End If
                                        End If
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initializeControls(oForm)
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

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_ORDR Then

            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try
            oApplication.Utilities.AddControls(oForm, "_53", "230", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, "", "Total Quantity", 0, 0, 0, False)
            oApplication.Utilities.AddControls(oForm, "_52", "222", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 0, 0, "", "", 0, 0, 0, True)
            oForm.Items.Item("_53").Visible = True
            oForm.Items.Item("_52").Visible = True
            oForm.Items.Item("_53").LinkTo = "_52"
            oForm.Items.Item("_52").RightJustified = True
            dataBind(oForm, "ORDR")
            oForm.Items.Item("_52").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub dataBind(ByVal oForm As SAPbouiCOM.Form, ByVal strTable As String)
        Try
            oEditText = oForm.Items.Item("_52").Specific
            oEditText.DataBind.SetBound(True, strTable, "U_DocQty")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function getDocumentTotal(ByVal oForm As SAPbouiCOM.Form) As Double
        Try
            Dim _retVal As Double = 0
            oMatrix = oForm.Items.Item("38").Specific
            For index As Integer = 1 To oMatrix.RowCount
                _retVal += oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "11", index))
                '_retVal += CDbl(CType(oMatrix.Columns.Item("11").Cells.Item(index).Specific, SAPbouiCOM.EditText).Value.ToString())
            Next
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

End Class
