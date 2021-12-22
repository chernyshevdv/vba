Attribute VB_Name = "mod_treeview"
    Option Compare Database

Public Sub load_tree(Optional a_select_node_key As String = "")
    Dim l_treeview As MSComctlLib.TreeView
    Dim l_node As MSComctlLib.Node
    Dim rs_cases As DAO.Recordset
    Dim l_key As String, l_icon As String
    
    Set l_treeview = Forms!frm_dashboard!tvw_cases.Object
    With l_treeview
        .LineStyle = tvwRootLines
        .HideSelection = False
        .SingleSel = True
        .Indentation = 200
        .Font.Size = 10
        Set .ImageList = Forms!frm_dashboard!imgl_icons.Object
    End With
    l_treeview.Nodes.Clear
    
    Set rs_cases = CurrentDb.OpenRecordset("SELECT * FROM cases WHERE status<>'archive'", dbOpenDynaset)
    While Not rs_cases.EOF
        l_key = "C_" & rs_cases!id
        l_icon = rs_cases("type")
        Set l_node = l_treeview.Nodes.Add(, , l_key, rs_cases("title"), l_icon)
        l_node.Bold = True
        load_tasks l_treeview, l_node, 0, rs_cases!id
        rs_cases.MoveNext
    Wend
    rs_cases.Close
    If a_select_node_key <> "" Then
        On Error Resume Next
        l_treeview.Nodes(a_select_node_key).Selected = True
        On Error GoTo 0
    End If
End Sub

Public Function get_node_ID(a_node As Node) As Long
    Dim l_key As String
    
    l_key = a_node.Key
    get_node_ID = CLng(Mid(l_key, 3))
End Function

Public Function get_node_type(a_node As Node) As String
    Dim l_key As String
    
    l_key = a_node.Key
    get_node_type = Left(l_key, 1)
End Function

Public Sub load_tasks(a_tvw As MSComctlLib.TreeView, a_parent_node As Node, a_parent_task_id As Long, a_case_id As Long)
    Dim l_rs As DAO.Recordset
    Dim l_node As Node, l_key As String, l_task_type As String
    
    Set l_rs = CurrentDb.OpenRecordset("SELECT * FROM tasks WHERE status<>'archive' AND case_id=" & a_case_id & " AND parent_id=" & a_parent_task_id & " ORDER BY priority", dbOpenDynaset)
    While Not l_rs.EOF
        l_key = "T_" & l_rs!id
        l_task_type = "task"
        If Not IsNull(l_rs!is_deliverable) And l_rs!is_deliverable Then l_task_type = "deliverable"
        If Not IsNull(l_rs!period) Then l_task_type = "periodic"
        Set l_node = a_tvw.Nodes.Add(a_parent_node, MSComctlLib.tvwChild, l_key, l_rs!title, l_task_type)
        If l_rs!status = "closed" Or l_rs!status = "archive" Then
            l_node.ForeColor = RGB(200, 200, 200)
        ElseIf l_rs!status = "pending" Then
            l_node.ForeColor = vbCyan
        ElseIf Not IsNull(l_rs!delegate_id) Then
            l_node.ForeColor = vbBlue
        End If
        load_jobs a_tvw, l_node, l_rs!id
        load_tasks a_tvw, l_node, l_rs!id, a_case_id
        l_rs.MoveNext
    Wend
    l_rs.Close
End Sub

Public Sub load_jobs(a_tvw As MSComctlLib.TreeView, a_parent_node As Node, a_parent_task_id As Long)
    Dim l_rs As DAO.Recordset
    Dim l_node As Node, l_key As String, l_sql As String
    
    l_sql = "SELECT * FROM jobs WHERE status<>'Archive' AND task_id=" & a_parent_task_id & " ORDER BY priority"
    Set l_rs = CurrentDb.OpenRecordset(l_sql, dbOpenDynaset)
    While Not l_rs.EOF
        l_key = "J_" & l_rs!id
        Set l_node = a_tvw.Nodes.Add(a_parent_node, MSComctlLib.tvwChild, l_key, l_rs!title, "job")
        If l_rs!status = "Closed" Then
            l_node.ForeColor = RGB(200, 200, 200)
        End If
        l_rs.MoveNext
    Wend
    l_rs.Close
End Sub

Public Sub go_to_task(task_id As Long)
    Dim l_treeview As MSComctlLib.TreeView
    
    mod_helpers.set_status_bar_text "Moving to task " & task_id
    Forms!frm_dashboard!tabDashboard.Pages!pg_cases.SetFocus
    Set l_treeview = Forms!frm_dashboard!tvw_cases.Object
    l_treeview.Nodes("T_" & task_id).Selected = True
End Sub
