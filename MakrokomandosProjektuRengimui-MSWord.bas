Sub A_PagamintiLyginamajiIsKeitimuSekimo()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Makrokomanda, skirta keitimø sekimams (track changes) paversti formatavimu, reikalingu teisës akto
' lyginamajame variante: keitimø sekimo áterptu tekstu paþymëtas (paprastai - pabrauktas) tekstas
' paverèiamas paryðkintu tekstu, keitimo sekimo iðbrauktu tekstu paþymëtas tekstas paþymimas perbraukimu,
' tekstas, paþymëtas ir áterpimo, ir iðbraukimo þymëmis, paþymimas já iðskiriant þydra spalva
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim chgAdd As Word.Revision
' informuojame vartotojà, jei dokumentas be susektø pakeitimø
If ActiveDocument.Revisions.Count = 0 Then
    MsgBox "Ðiame dokumente nëra uþfiksuotø pakeitimø", vbOKOnly + vbInformation
Else
    ActiveDocument.TrackRevisions = False
    ' ciklas visø pakeitimø perþiûrai
    For Each chgAdd In ActiveDocument.Revisions
        ' keièiam susektus iðbraukimus á paprastà iðbrauktà tekstà
        If chgAdd.Type = wdRevisionDelete Then
            chgAdd.Range.Font.StrikeThrough = True
            chgAdd.Reject
        ' perspëjam vartotojà, jei aptiktas teksto perkëlimas;
        ' tokius makrokomanda palieka nepakeistus, tik paþymi
        ElseIf chgAdd.Type = wdRevisionMovedFrom Then
            MsgBox ("Makrokomanda nepalaiko teksto perkëlimø (tik trynimà/áterpimà)."), vbOKOnly + vbExclamation
            chgAdd.Range.Select ' move insertion point
        ElseIf chgAdd.Type = wdRevisionMovedTo Then
            MsgBox ("Makrokomanda nepalaiko teksto perkëlimø (tik trynimà/áterpimà)."), vbOKOnly + vbExclamation
            chgAdd.Range.Select ' move insertion point
        ' keièiam susektus áterpimus á paryðkintà tekstà
        ElseIf chgAdd.Type = wdRevisionInsert Then
            chgAdd.Range.Font.Bold = True
            chgAdd.Accept
        ' bet kokius kitokius pakeitimus paþymime þaliai ir perspëjame vartotojà
        Else
            MsgBox ("Rastas kitoks teksto pakeitimas: jis priimtas ir paþymëtas þalsvai."), vbOKOnly + vbExclamation
            chgAdd.Range.HighlightColorIndex = wdBrightGreen
            chgAdd.Accept
            ' chgAdd.Range.Select ' move insertion point
        End If
    Next chgAdd
End If
' makrokomandos dalis paþymëti tekstui, kuris yra ir paryðkintas, ir iðbrauktas;
' taip nutinka, kai dokumentas buvo taisomas dviejø skirtingø autoriø, kuriø vienas áraðo pakeitimus,
' o kitas juos ar dalá jø panaikina
MsgBox ("Jei bus rasta konfliktuojanèiø dviejø autoriø keitimø, jie bus paþymëti þydrai"), vbOKOnly + vbInformation
    Options.DefaultHighlightColorIndex = wdTurquoise
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .StrikeThrough = True
        .DoubleStrikeThrough = False
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = wdPink ' ið tiesø veikia aukðèiau esanti Option... parinktis spalvai parinkti
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
' Makrokomanda, ið lyginamojo varianto padaranti galutiná teisës akto variantà; prieð pritaikant ðià
' makro komandà reikia ásitikinti, kad tekstas, kurio nenorime iðtrinti (pvz., straipsniø pavadinimai),
' yra be paryðkinimo (visi paryðkinti teksto elementai bus iðtrinami)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CarryOn = MsgBox("Ði makrokomanda iðtrins visà iðbrauktà tekstà. Jà leiskite tik ásitikinæ, kad teksto dalys, kuriø iðtrinti nenorite, nëra iðbrauktos. Ar norite tæsti?", vbYesNo, "Ar tikrai to norite?")
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
' Makrokomanda, aptinkanti ir paþyminti klaidas, kai lyginamajame variante tas pats tekstas yra paþymëtas
' ir kaip áterptas (paryðkintas), ir kaip iðbrauktas (pritaikius iðbraukimo formatavimà)
' Makrokomanda paþymi tokias teksto vietas
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
' Makrokomanda, aptinkanti ir iðtrinanti klaidas, kai lyginamajame variante tas pats tekstas yra paþymëtas
' ir kaip áterptas (paryðkintas), ir kaip iðbrauktas (pritaikius iðbraukimo formatavimà)
' Makrokomanda iðtrina tokias teksto vietas. Jà siûlytina taikyti tik po to, kai tokios vietos yra
' jau suþymëtos makrokomandos „C_PazymetiParyskintaIrIsbrauktaTeksta()“ ir vartotojas ásitikino, kad
' neiðtrins nieko naudingo
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
' Su ðia makrokomanda galima ið lyginamojo teisës akto varianto pagaminti „pirminá“ variantà
' Makrokomanda paprasèiausiai iðtrina paryðkintà tekstà.
' Ávykdæ ðià makrokomandà galite gautà rezultatà palyginti (naudojant dokumentø lyginimo funkcijà) su
' pirminiu keièiamo teisës akto variantu siekdami ásitikinti, kad nepadarëte klaidø sugadindami pirminá
' tekstà.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CarryOn = MsgBox("Ði makrokomanda iðtrins visà paryðkintà tekstà. Jà leiskite tik ásitikinæ, kad teksto dalys, kuriø iðtrinti nenorite, nëra paryðkintos. Ar norite tæsti?", vbYesNo, "Ar tikrai to norite?")
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
