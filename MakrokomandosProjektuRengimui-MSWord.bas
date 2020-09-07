' Makrokomandų koduotė Github'e - UTF8. Norint  makrokomandas naudoti MS Word tekstą reikia konvertuoti
' į CP1257

Sub A_PagamintiLyginamajiIsKeitimuSekimo()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Makrokomanda, skirta keitimų sekimams (track changes) paversti formatavimu, reikalingu teisės akto
' lyginamajame variante: keitimų sekimo įterptu tekstu pažymėtas (paprastai - pabrauktas) tekstas
' paverčiamas paryškintu tekstu, keitimo sekimo išbrauktu tekstu pažymėtas tekstas pažymimas perbraukimu,
' tekstas, pažymėtas ir įterpimo, ir išbraukimo žymėmis, pažymimas jį išskiriant žydra spalva
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim chgAdd As Word.Revision
' informuojame vartotoją, jei dokumentas be susektų pakeitimų
If ActiveDocument.Revisions.Count = 0 Then
    MsgBox "Šiame dokumente nėra užfiksuotų pakeitimų", vbOKOnly + vbInformation
Else
    ActiveDocument.TrackRevisions = False
    ' ciklas visų pakeitimų peržiūrai
    For Each chgAdd In ActiveDocument.Revisions
        ' keičiam susektus išbraukimus į paprastą išbrauktą tekstą
        If chgAdd.Type = wdRevisionDelete Then
            chgAdd.Range.Font.StrikeThrough = True
            chgAdd.Reject
        ' perspėjam vartotoją, jei aptiktas teksto perkėlimas;
        ' tokius makrokomanda palieka nepakeistus, tik pažymi
        ElseIf chgAdd.Type = wdRevisionMovedFrom Then
            MsgBox ("Makrokomanda nepalaiko teksto perkėlimų (tik trynimą/įterpimą)."), vbOKOnly + vbExclamation
            chgAdd.Range.Select ' move insertion point
        ElseIf chgAdd.Type = wdRevisionMovedTo Then
            MsgBox ("Makrokomanda nepalaiko teksto perkėlimų (tik trynimą/įterpimą)."), vbOKOnly + vbExclamation
            chgAdd.Range.Select ' move insertion point
        ' keičiam susektus įterpimus į paryškintą tekstą
        ElseIf chgAdd.Type = wdRevisionInsert Then
            chgAdd.Range.Font.Bold = True
            chgAdd.Accept
        ' bet kokius kitokius pakeitimus pažymime žaliai ir perspėjame vartotoją
        Else
            MsgBox ("Rastas kitoks teksto pakeitimas: jis priimtas ir pažymėtas žalsvai."), vbOKOnly + vbExclamation
            chgAdd.Range.HighlightColorIndex = wdBrightGreen
            chgAdd.Accept
            ' chgAdd.Range.Select ' move insertion point
        End If
    Next chgAdd
End If
' makrokomandos dalis pažymėti tekstui, kuris yra ir paryškintas, ir išbrauktas;
' taip nutinka, kai dokumentas buvo taisomas dviejų skirtingų autorių, kurių vienas įrašo pakeitimus,
' o kitas juos ar dalį jų panaikina
MsgBox ("Jei bus rasta konfliktuojančių dviejų autorių keitimų, jie bus pažymėti žydrai"), vbOKOnly + vbInformation
    Options.DefaultHighlightColorIndex = wdTurquoise
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .StrikeThrough = True
        .DoubleStrikeThrough = False
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = wdPink ' iš tiesų veikia aukščiau esanti Option... parinktis spalvai parinkti
    With Selection.Find
        .Text = ""
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub



Sub B_PagamintiGalutiniIsLyginamojo()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Makrokomanda, iš lyginamojo varianto padaranti galutinį teisės akto variantą; prieš pritaikant šią
' makro komandą reikia įsitikinti, kad tekstas, kurio nenorime ištrinti (pvz., straipsnių pavadinimai),
' yra be paryškinimo (visi paryškinti teksto elementai bus ištrinami)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CarryOn = MsgBox("Ši makrokomanda ištrins visą išbrauktą tekstą. Ją leiskite tik įsitikinę, kad teksto dalys, kurių ištrinti nenorite, nėra išbrauktos. Ar norite tęsti?", vbYesNo, "Ar tikrai to norite?")
If CarryOn = vbYes Then

    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .StrikeThrough = True
        .DoubleStrikeThrough = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End If
End Sub



Sub C_PazymetiParyskintaIrIsbrauktaTeksta()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Makrokomanda, aptinkanti ir pažyminti klaidas, kai lyginamajame variante tas pats tekstas yra pažymėtas
' ir kaip įterptas (paryškintas), ir kaip išbrauktas (pritaikius išbraukimo formatavimą)
' Makrokomanda pažymi tokias teksto vietas
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .StrikeThrough = True
        .DoubleStrikeThrough = False
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = wdYellow
    

    With Selection.Find
        .Text = ""
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub



Sub D_IstrintiParyskintaIrIsbrauktaTeksta()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Makrokomanda, aptinkanti ir ištrinanti klaidas, kai lyginamajame variante tas pats tekstas yra pažymėtas
' ir kaip įterptas (paryškintas), ir kaip išbrauktas (pritaikius išbraukimo formatavimą)
' Makrokomanda ištrina tokias teksto vietas. Ją siūlytina taikyti tik po to, kai tokios vietos yra
' jau sužymėtos makrokomandos „C_PazymetiParyskintaIrIsbrauktaTeksta()“ ir vartotojas įsitikino, kad
' neištrins nieko naudingo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .StrikeThrough = True
        .DoubleStrikeThrough = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub



Sub E_PagamintiPirminiIsLyginamojo()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Su šia makrokomanda galima iš lyginamojo teisės akto varianto pagaminti „pirminį“ variantą
' Makrokomanda paprasčiausiai ištrina paryškintą tekstą.
' Įvykdę šią makrokomandą galite gautą rezultatą palyginti (naudojant dokumentų lyginimo funkciją) su
' pirminiu keičiamo teisės akto variantu siekdami įsitikinti, kad nepadarėte klaidų sugadindami pirminį
' tekstą.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CarryOn = MsgBox("Ši makrokomanda ištrins visą paryškintą tekstą. Ją leiskite tik įsitikinę, kad teksto dalys, kurių ištrinti nenorite, nėra paryškintos. Ar norite tęsti?", vbYesNo, "Ar tikrai to norite?")
If CarryOn = vbYes Then

    Selection.Find.ClearFormatting
    Selection.Find.Font.Bold = True
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End If
End Sub
