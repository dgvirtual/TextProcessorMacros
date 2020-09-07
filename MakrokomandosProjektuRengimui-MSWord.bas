Sub A_PagamintiLyginamajiIsKeitimuSekimo()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Makrokomanda, skirta keitim� sekimams (track changes) paversti formatavimu, reikalingu teis�s akto
' lyginamajame variante: keitim� sekimo �terptu tekstu pa�ym�tas (paprastai - pabrauktas) tekstas
' paver�iamas pary�kintu tekstu, keitimo sekimo i�brauktu tekstu pa�ym�tas tekstas pa�ymimas perbraukimu,
' tekstas, pa�ym�tas ir �terpimo, ir i�braukimo �ym�mis, pa�ymimas j� i�skiriant �ydra spalva
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim chgAdd As Word.Revision
' informuojame vartotoj�, jei dokumentas be susekt� pakeitim�
If ActiveDocument.Revisions.Count = 0 Then
    MsgBox "�iame dokumente n�ra u�fiksuot� pakeitim�", vbOKOnly + vbInformation
Else
    ActiveDocument.TrackRevisions = False
    ' ciklas vis� pakeitim� per�i�rai
    For Each chgAdd In ActiveDocument.Revisions
        ' kei�iam susektus i�braukimus � paprast� i�braukt� tekst�
        If chgAdd.Type = wdRevisionDelete Then
            chgAdd.Range.Font.StrikeThrough = True
            chgAdd.Reject
        ' persp�jam vartotoj�, jei aptiktas teksto perk�limas;
        ' tokius makrokomanda palieka nepakeistus, tik pa�ymi
        ElseIf chgAdd.Type = wdRevisionMovedFrom Then
            MsgBox ("Makrokomanda nepalaiko teksto perk�lim� (tik trynim�/�terpim�)."), vbOKOnly + vbExclamation
            chgAdd.Range.Select ' move insertion point
        ElseIf chgAdd.Type = wdRevisionMovedTo Then
            MsgBox ("Makrokomanda nepalaiko teksto perk�lim� (tik trynim�/�terpim�)."), vbOKOnly + vbExclamation
            chgAdd.Range.Select ' move insertion point
        ' kei�iam susektus �terpimus � pary�kint� tekst�
        ElseIf chgAdd.Type = wdRevisionInsert Then
            chgAdd.Range.Font.Bold = True
            chgAdd.Accept
        ' bet kokius kitokius pakeitimus pa�ymime �aliai ir persp�jame vartotoj�
        Else
            MsgBox ("Rastas kitoks teksto pakeitimas: jis priimtas ir pa�ym�tas �alsvai."), vbOKOnly + vbExclamation
            chgAdd.Range.HighlightColorIndex = wdBrightGreen
            chgAdd.Accept
            ' chgAdd.Range.Select ' move insertion point
        End If
    Next chgAdd
End If
' makrokomandos dalis pa�ym�ti tekstui, kuris yra ir pary�kintas, ir i�brauktas;
' taip nutinka, kai dokumentas buvo taisomas dviej� skirting� autori�, kuri� vienas �ra�o pakeitimus,
' o kitas juos ar dal� j� panaikina
MsgBox ("Jei bus rasta konfliktuojan�i� dviej� autori� keitim�, jie bus pa�ym�ti �ydrai"), vbOKOnly + vbInformation
    Options.DefaultHighlightColorIndex = wdTurquoise
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .StrikeThrough = True
        .DoubleStrikeThrough = False
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = wdPink ' i� ties� veikia auk��iau esanti Option... parinktis spalvai parinkti
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
' Makrokomanda, i� lyginamojo varianto padaranti galutin� teis�s akto variant�; prie� pritaikant �i�
' makro komand� reikia �sitikinti, kad tekstas, kurio nenorime i�trinti (pvz., straipsni� pavadinimai),
' yra be pary�kinimo (visi pary�kinti teksto elementai bus i�trinami)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CarryOn = MsgBox("�i makrokomanda i�trins vis� i�braukt� tekst�. J� leiskite tik �sitikin�, kad teksto dalys, kuri� i�trinti nenorite, n�ra i�brauktos. Ar norite t�sti?", vbYesNo, "Ar tikrai to norite?")
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
' Makrokomanda, aptinkanti ir pa�yminti klaidas, kai lyginamajame variante tas pats tekstas yra pa�ym�tas
' ir kaip �terptas (pary�kintas), ir kaip i�brauktas (pritaikius i�braukimo formatavim�)
' Makrokomanda pa�ymi tokias teksto vietas
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
' Makrokomanda, aptinkanti ir i�trinanti klaidas, kai lyginamajame variante tas pats tekstas yra pa�ym�tas
' ir kaip �terptas (pary�kintas), ir kaip i�brauktas (pritaikius i�braukimo formatavim�)
' Makrokomanda i�trina tokias teksto vietas. J� si�lytina taikyti tik po to, kai tokios vietos yra
' jau su�ym�tos makrokomandos �C_PazymetiParyskintaIrIsbrauktaTeksta()� ir vartotojas �sitikino, kad
' nei�trins nieko naudingo
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
' Su �ia makrokomanda galima i� lyginamojo teis�s akto varianto pagaminti �pirmin� variant�
' Makrokomanda papras�iausiai i�trina pary�kint� tekst�.
' �vykd� �i� makrokomand� galite gaut� rezultat� palyginti (naudojant dokument� lyginimo funkcij�) su
' pirminiu kei�iamo teis�s akto variantu siekdami �sitikinti, kad nepadar�te klaid� sugadindami pirmin�
' tekst�.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CarryOn = MsgBox("�i makrokomanda i�trins vis� pary�kint� tekst�. J� leiskite tik �sitikin�, kad teksto dalys, kuri� i�trinti nenorite, n�ra pary�kintos. Ar norite t�sti?", vbYesNo, "Ar tikrai to norite?")
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
