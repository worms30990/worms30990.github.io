Attribute VB_Name = "ESKD"
Sub NewMacros1()
Attribute NewMacros1.VB_Description = "Установлки стандартів Єдиної конструкторської документації "
Attribute NewMacros1.VB_ProcData.VB_Invoke_Func = "Normal.ESCD.ЄСКД"

    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 14
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    ActiveDocument.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5
    With Selection.PageSetup
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
    End With
End Sub
