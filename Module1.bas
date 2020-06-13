Attribute VB_Name = "Module1"
Sub main()
    'Application.WindowState = xlMinimized
    With UserForm1
        .TextBox2.Locked = True
        .TextBox3.Locked = True
        .TextBox4.Locked = True
        .TextBox5.Locked = True
        .Label1.Font.Bold = True
        .Label3.Font.Bold = True
        .Show
    End With
    'Application.WindowState = xlMaximized
End Sub

Sub R_add_item()
    If IsNumeric(UserForm1.TextBox1.Value) Then
        UserForm1.ListBox1.AddItem (UserForm1.TextBox1.Value)
        Module1.R_vktr_size
        Module1.DotProduct
        Module1.AdjProj
        Module1.VktrProj
    Else
        MsgBox "Is this a numeric entry?"
    End If
End Sub

Sub R_remove_item()
    If UserForm1.ListBox1.ListIndex >= 0 Then
        UserForm1.ListBox1.RemoveItem (UserForm1.ListBox1.ListIndex)
        Module1.R_vktr_size
        Module1.DotProduct
        Module1.AdjProj
        Module1.VktrProj
    Else
        UserForm1.ListBox1.Selected(UserForm1.ListBox1.ListCount - 1) = True
    End If
End Sub

Sub R_vktr_size()
    Dim SumSqrs As Double
    SumSqrs = 0
    
    For i = 0 To UserForm1.ListBox1.ListCount - 1
        SumSqrs = SumSqrs + (CDbl(UserForm1.ListBox1.List(i)) * CDbl(UserForm1.ListBox1.List(i)))
    Next

    UserForm1.TextBox2.Value = Sqr(SumSqrs)
    
End Sub

Sub S_add_item()
    If IsNumeric(UserForm1.TextBox1.Value) Then
        UserForm1.ListBox2.AddItem (UserForm1.TextBox1.Value)
        Module1.S_vktr_size
        Module1.DotProduct
        Module1.AdjProj
        Module1.VktrProj
    Else
        MsgBox "Is this a numeric entry?"
    End If
End Sub

Sub S_remove_item()
    If UserForm1.ListBox2.ListIndex >= 0 Then
        UserForm1.ListBox2.RemoveItem (UserForm1.ListBox2.ListIndex)
        Module1.S_vktr_size
        Module1.DotProduct
        Module1.AdjProj
        Module1.VktrProj
    Else
        UserForm1.ListBox2.Selected(UserForm1.ListBox2.ListCount - 1) = True
    End If
End Sub

Sub S_vktr_size()
    Dim SumSqrs As Double
    SumSqrs = 0
    
    For i = 0 To UserForm1.ListBox2.ListCount - 1
        SumSqrs = SumSqrs + (CDbl(UserForm1.ListBox2.List(i)) * CDbl(UserForm1.ListBox2.List(i)))
    Next

    UserForm1.TextBox3.Value = Sqr(SumSqrs)
    
End Sub

Sub DotProduct()
    Dim DotPrdkt As Double
    DotPrdkt = 0
    If UserForm1.ListBox1.ListCount = UserForm1.ListBox2.ListCount Then
        For i = 0 To UserForm1.ListBox1.ListCount - 1
            DotPrdkt = DotPrdkt + (CDbl(UserForm1.ListBox1.List(i)) * CDbl(UserForm1.ListBox2.List(i)))
        Next
        UserForm1.TextBox4.Value = DotPrdkt
    Else
        UserForm1.TextBox4.Value = "N/A"
    End If
End Sub

Sub AdjProj()
    If IsNumeric(UserForm1.TextBox4.Value) Then
        UserForm1.TextBox5.Value = CDbl(UserForm1.TextBox4.Value) / CDbl(UserForm1.TextBox2.Value)
    Else
        UserForm1.TextBox5.Value = "N/A"
    End If
End Sub

Sub VktrProj()
    Dim Skle As Double
    UserForm1.ListBox3.Clear
    If IsNumeric(UserForm1.TextBox4.Value) Then
        Skle = (CDbl(UserForm1.TextBox4.Value) / (CDbl(UserForm1.TextBox2.Value) * CDbl(UserForm1.TextBox2.Value)))
        For i = 0 To UserForm1.ListBox1.ListCount - 1
            UserForm1.ListBox3.AddItem (UserForm1.ListBox1.List(i) * Skle)
        Next
    Else: End If
End Sub

Sub clear_all()
    With UserForm1
        .ListBox1.Clear
        .ListBox2.Clear
        .ListBox3.Clear
        .TextBox1.Value = ""
        .TextBox2.Value = ""
        .TextBox3.Value = ""
        .TextBox4.Value = ""
        .TextBox5.Value = ""
    End With
End Sub
