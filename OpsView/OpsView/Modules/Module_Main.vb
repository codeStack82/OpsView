Imports System.Text
Imports System.Windows.Forms

Module Module_Main
    Public wellCounter As Integer = 0

    'Main Method
    Sub Main()

        Dim districts As Array = linq_getDistrictNames()
        Dim hash_ActiveWells As Hashtable

        hash_ActiveWells = qry_getWVActiveWells()
        'list_HistoricWells = qry_getWVHistoricWells()

        'Create new Main form
        Dim frm_Main = New frm_OpsView

        Dim key As ICollection = hash_ActiveWells.Keys

        'create
        Dim wellCount = hash_ActiveWells.Keys.Count
        Dim wellList(wellCount) As String

        'Add wells to str array
        Dim i As Integer = 1
        For Each k In key

            Dim temp As String = getDistAbbvr(hash_ActiveWells(k).ToString())
            Dim dist As String = temp
            Dim well As String = k.ToString()

            wellList(i) = dist & " - " & well

            i += 1
        Next

        Array.Sort(wellList)


        For Each w In wellList
            Console.WriteLine(w)
        Next

        Dim x As Integer = 1
        For Each k In key

            Dim wellInfo As String = wellList(x)

            'Create new well panel
            Dim pnl_WellPanel = WellPanel(wellInfo)

            'Set Panel in flow Panel
            frm_Main.fpnl_FlowPanel.Controls.Add(pnl_WellPanel)

            x += 1
        Next

        'Show Form
        frm_Main.ShowDialog()
    End Sub

    'Func to createa new well panel
    Function WellPanel(ByVal wellList As String) As Panel

        'Declare new panel components
        Dim pnl_MainWellPanel As New Panel
        Dim pnl_TopWellPanel As New Panel
        Dim pnl_BtmWellPanel As New Panel

        'Declare label component(s)
        Dim lbl_TopWellName As New Label

        'Increment well panel count
        wellCounter += 1

        'Set panel properties 
        With pnl_MainWellPanel
            .Name = "pnl_Main_" & wellList
            .Width = 400
            .Height = 400
            .BackColor = Color.GhostWhite
            .Margin = New Padding(5, 5, 5, 5)
            .BorderStyle = BorderStyle.FixedSingle
        End With

        With pnl_TopWellPanel
            .Name = "pnl_Top_" & wellList
            .Dock = DockStyle.Top
            .Height = 40
            .BackColor = Color.Gainsboro
            .Margin = New Padding(5, 5, 5, 5)
            .Padding = New Padding(10, 5, 5, 5)
            .BorderStyle = BorderStyle.None
        End With

        With pnl_BtmWellPanel
            .Name = "pnl_Btm_" & wellList
            .Dock = DockStyle.Bottom
            .Height = 35
            .BackColor = Color.Gainsboro
            .Margin = New Padding(5, 5, 5, 5)
            .Padding = New Padding(5, 5, 5, 5)
            .BorderStyle = BorderStyle.None
        End With

        'Set label properties
        With lbl_TopWellName
            .Name = "lbl_Top_" & wellList
            .Text = wellList
            .Width = 400
            .BorderStyle = BorderStyle.None
            .Font = New Drawing.Font("Times New Roman", 11, FontStyle.Bold)
        End With

        'Add conponents to main well panel~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        'Add Well name label to Top panel
        pnl_TopWellPanel.Controls.Add(lbl_TopWellName)

        'Add Top panel to main panel
        pnl_MainWellPanel.Controls.Add(pnl_TopWellPanel)

        'Add Btm panel to main panel
        pnl_MainWellPanel.Controls.Add(pnl_BtmWellPanel)

        'Return well panel
        Return pnl_MainWellPanel
    End Function

    Function getDistAbbvr(ByVal dist As String) As String
        Dim District As String

        If dist = "EAGLE FORD" Then
            District = "EGFD"
        ElseIf dist = "EASTERN GULF COAST" Then
            District = "EGC"
        ElseIf dist = "MID-CONTINENT" Then
            District = "MCON"
        ElseIf dist = "POWDER RIVER" Then
            District = "PDR"
        ElseIf dist = "MARCELLUS NORTH" Then
            District = "MRCN"
        ElseIf dist = "MARCELLUS SOUTH" Then
            District = "MRCNS"
        ElseIf dist = "UTICA" Then
            District = "UTC"
        Else
            District = dist
        End If

        Return District
    End Function
End Module
