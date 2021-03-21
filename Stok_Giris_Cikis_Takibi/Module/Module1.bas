Attribute VB_Name = "Module1"
Sub Kayýt()
Attribute Kayýt.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Kayýt Makro
'

'
    ActiveCell.FormulaR1C1 = ""
    Range("A2:G2").Select
    Selection.Copy
    Sheets("STOK HAREKETLERÝ").Select
    Application.Goto Reference:="R65000C1"
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range("A8").Select
    ActiveCell.Select
    Range("A7").Select
    Sheets("GÝRÝÞ-ÇIKIÞ").Select
    Range("A4").Select
End Sub
Sub Makro2()
Attribute Makro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro2 Makro
'

'
    Sheets("Kayýt").Select
    Range("A2:G2").Select
    Selection.Copy
    Sheets("STOK HAREKETLERÝ").Select
    Application.Goto Reference:="R65000C1"
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, 0).Range("A1").Select
    Sheets("GÝRÝÞ-ÇIKIÞ").Select
    Range("A4").Select
End Sub
Sub listele()
Attribute listele.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' listele Makro
'
' Klavye Kýsayolu: Ctrl+d
'
End Sub
Sub Makro4()
Attribute Makro4.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' Makro4 Makro
'
' Klavye Kýsayolu: Ctrl+s
'
    Range("C14").Select
    Range("ListeAdý").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range _
        ("Kriter"), CopyToRange:=Range("C13:I13"), Unique:=False
    ActiveWindow.SmallScroll Down:=11
End Sub
Sub ListeleGrup()
Attribute ListeleGrup.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ListeleGrup Makro
'

'
    Range("D4").Select
    Selection.Copy
    Range("A2").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B2:G2").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Application.Run "'Stok Giriþ-Çýkýþ Takibi.xlsm'!Makro4"
    ActiveWindow.SmallScroll Down:=-11
End Sub
Sub ListeleKod()
Attribute ListeleKod.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ListeleKod Makro
'

'
    Range("D5").Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C2:G2").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Application.Run "'Stok Giriþ-Çýkýþ Takibi.xlsm'!Makro4"
    ActiveWindow.SmallScroll Down:=-33
End Sub
Sub Makro8()
Attribute Makro8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro8 Makro
'

'
    ActiveCell.FormulaR1C1 = ""
    Range("D5").Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C2:G2").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Application.Run "'Stok Giriþ-Çýkýþ Takibi.xlsm'!Makro4"
    ActiveWindow.SmallScroll Down:=-11
End Sub
