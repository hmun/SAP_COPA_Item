VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCOPAActuals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsConnected As Integer

Private Sub Class_Initialize()
    IsConnected = SAPCheck()
End Sub

Public Function PostCostingBasedData(pOperatingConcern As String, pData As Collection) As String
    On Error GoTo PostCostingBasedData_Error
    Dim lSAPFormat As New SAPFormat
    Dim boCOPAActuals As Object
    Dim lOperatingConcern
    Dim oIndexstructure As Object
    Dim oCoobject As Object
    Dim oPervalue As Object
    Dim oRETURN As Object
    Set boCOPAActuals = MySAPCon.SAPBapiControl.GetSAPObject("COPAActuals")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boCOPAActuals, "PostCostingBasedData", "Return")
    Set oInputData = MySAPCon.SAPBapiControl.DimAs(boCOPAActuals, "PostCostingBasedData", "InputData")
    Set oFieldList = MySAPCon.SAPBapiControl.DimAs(boCOPAActuals, "PostCostingBasedData", "FieldList")
    lOperatingConcern = pOperatingConcern

    Dim aIDRow As Object
    Dim aFLRow As Object
    Dim aRow As Object
    Dim aItem As Object
    lCnt = 0
    For Each aRow In pData
        lCnt = lCnt + 1
        For Each aItem In aRow
            Set aIDRow = oInputData.Rows.Add
            aIDRow.Value("RECORD_ID") = lCnt
            aIDRow.Value("FIELDNAME") = aItem.gFIELDNAME
            aIDRow.Value("VALUE") = aItem.gVALUE
            If Not IsNull(aItem.gCURRENCY) And aItem.gCURRENCY <> "" Then
                aIDRow.Value("CURRENCY") = aItem.gCURRENCY
            End If
            If lCnt = 1 Then
                Set aFLRow = oFieldList.Rows.Add
                aFLRow.Value("FIELDNAME") = aItem.gFIELDNAME
            End If
        Next aItem
    Next aRow
    Application.Cursor = xlWait
    boCOPAActuals.PostCostingBasedData OperatingConcern:=lOperatingConcern, _
    TestRun:="", _
    InputData:=oInputData, _
    FieldList:=oFieldList, _
    Return:=oRETURN
    Application.Cursor = xlDefault
    If oRETURN.RowCount = 0 Then
        PostCostingBasedData = "Success"
        Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
        aSAPBapiTranctionCommit.commit
    Else
        Dim RetRow As Object
        For Each RetRow In oRETURN.Rows
            PostCostingBasedData = PostCostingBasedData & ";" & RetRow("MESSAGE")
        Next RetRow
    End If
    Set boCostActivityPlanning = Nothing
    Exit Function
PostCostingBasedData_Error:
    Application.Cursor = xlDefault
    PostCostingBasedData = "Error"
    MySAPErr.MSGProt "SAPCOPAActuals", "PostCostingBasedData", "", Err.Number, Err.Description
    Exit Function
End Function

