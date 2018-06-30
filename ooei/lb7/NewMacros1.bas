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
    With Selection.Find
        .Text = "^s"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdStore
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=10, NumColumns:= _
        2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    Selection.Tables(1).Select
    Selection.SelectCell
End Sub
