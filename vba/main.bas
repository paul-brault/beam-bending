Attribute VB_Name = "main"
Sub Call_main()
Dim t As Single
t = Timer

Debug.Print vbNewLine
Debug.Print Time

data = Sheets("input").Range("A1").Value
data = NormalizeDataLocale(data) 'dot or comma for decimal
result_compliance = compliance(data)

If result_compliance = "ok" Then
    Call main(data)
    'MsgBox("Texte", vbYesNoCancel + vbExclamation + vbDefaultButton2, "Titre")
    MsgBox Round(Timer - t, 3) * 1000 & "ms" ', vbInformation, "Succès")
Else
    MsgBox (result_compliance)
End If


End Sub
Function main(data As Variant)

    For i = 0 To 100
        Debug.Print ""
    Next i
data = Sheets("input").Range("A1").Value 'useless
data = NormalizeDataLocale(data) 'useless

'Dim noeud__axe() 'm
Dim noeud__appui() 'Bool
Dim noeud__fext() 'N
Dim noeud__ry()
Dim noeud__mz()

Dim element__long() 'm
Dim element__young() 'N/m^2
Dim element__iz() 'm^4
Dim element__qext() 'N/m
Dim element__k() 'tenseur rigidité (matrice * element)
Dim element__f() 'matrice force (vecteur * element)

Dim dl__k() 'matrice de rigidite globale
Dim dl__f() 'vecteur force global
Dim dl__u() 'vecteur déplacement global

Dim dle__feq(3) 'vecteur force équivalente élémentaire

Dim x__uy()
Dim x__ty()
Dim x__mfz()

Dim travee__uy_max() 'matrice flèche max // abscisse
Dim travee__mfz_0() 'matrice moment max // abscisse
Dim travee__mfz_pos_max()
Dim travee__mfz_neg_min()

'1 - MAILLAGE

data_split = Split(data, ";") '0axe_appui;1extremite_poutre;2young_poutre;3iz_poutre;4axe_ponctuelle;5force_ponctuelle;6origine_linéaire;7extremite_linéaire;8force_linéaire

noeud__axe = Split("0:" & data_split(0) & ":" & data_split(1) & ":" & data_split(4) & ":" & data_split(6) & ":" & data_split(7), ":")
Call ArraySort(noeud__axe, 0, UBound(noeud__axe))
Call ArrayUnique(noeud__axe)

nb_noeud = UBound(noeud__axe) + 1
nb_dl = nb_noeud * 2 'dl_element = i_element * 2 + k [k 0:3] // dl_noeud = i_noeud * 2 + k [k 0:1]
nb_element = nb_noeud - 1

ReDim noeud__appui(nb_noeud - 1)
ReDim noeud__fext(nb_noeud - 1)
axe_appuis = Split(data_split(0), ":")
axe_ponctuelles = Split(data_split(4), ":")
force_ponctuelles = Split(data_split(5), ":")
For i_noeud = 0 To nb_noeud - 1
    noeud__appui(i_noeud) = False
    For i_appui = 0 To UBound(axe_appuis)
        If noeud__axe(i_noeud) = axe_appuis(i_appui) Then noeud__appui(i_noeud) = True
    Next i_appui
    'Debug.Print noeud__appui(i_noeud)
    noeud__fext(i_noeud) = 0
    For i_ponctuelle = 0 To UBound(force_ponctuelles)
        If noeud__axe(i_noeud) = axe_ponctuelles(i_ponctuelle) Then noeud__fext(i_noeud) = noeud__fext(i_noeud) + force_ponctuelles(i_ponctuelle)
    Next i_ponctuelle
    'Debug.Print noeud__fext(i_noeud)
Next i_noeud

ReDim element__long(nb_element - 1)
ReDim element__young(nb_element - 1)
ReDim element__iz(nb_element - 1)
ReDim element__qext(nb_element - 1)
i_poutre = 0
extremite_poutres = Split(data_split(1), ":")
young_poutres = Split(data_split(2), ":")
iz_poutres = Split(data_split(3), ":")
origine_lineaires = Split(data_split(6), ":")
extremite_lineaires = Split(data_split(7), ":")
force_lineaires = Split(data_split(8), ":")
For i_element = 0 To nb_element - 1 'pour chaque element attribuer la longueur long, le module de young et le moment quadratique Iz
    element__long(i_element) = Round(noeud__axe(i_element + 1) - noeud__axe(i_element), 5)
    'Debug.Print i_poutre
    If Round(noeud__axe(i_element + 1), 5) > Round(extremite_poutres(i_poutre), 5) Then i_poutre = i_poutre + 1 ': Debug.Print noeud__axe(i_element + 1), extremite_poutres(i_poutre - 1)   'changer de poutre si on dépasse l'extremité
    element__young(i_element) = young_poutres(i_poutre)
    element__iz(i_element) = iz_poutres(i_poutre)
    element__qext(i_element) = 0
    For i_lineaire = 0 To UBound(force_lineaires)
        If Round(origine_lineaires(i_lineaire), 10) <= Round(noeud__axe(i_element), 10) And Round(extremite_lineaires(i_lineaire), 10) >= Round(noeud__axe(i_element + 1), 10) Then element__qext(i_element) = element__qext(i_element) + force_lineaires(i_lineaire)
    Next i_lineaire
    'Debug.Print element__qext(i_element)
Next i_element


''Afficher le maillage
'Debug.Print "1- Maillage"
'Debug.Print "Axe noeud", "Appui", "Fext"
'For i = LBound(noeud__axe) To UBound(noeud__axe)
'    Debug.Print noeud__axe(i), noeud__appui(i), noeud__fext(i)
'Next i
'Debug.Print "Long", "Young", "Iz", "Qext"
'For i = LBound(element__long) To UBound(element__long)
'    Debug.Print element__long(i), element__young(i), element__iz(i), element__qext(i)
'Next i
'End

'2 - MATRICE DE RIGIDITE GLOBALE
ReDim dl__k(nb_dl - 1, nb_dl - 1)
ReDim element__k(nb_element - 1, 3, 3)
For i_element = 0 To nb_element - 1
    constante = (element__young(i_element) * element__iz(i_element)) / (element__long(i_element) ^ 3)

    element__k(i_element, 0, 0) = constante * 12: dl__k(2 * i_element + 0, 0 + i_element * 2) = element__k(i_element, 0, 0) + dl__k(2 * i_element + 0, 0 + i_element * 2)
    element__k(i_element, 0, 1) = constante * element__long(i_element) * 6: dl__k(2 * i_element + 0, 1 + i_element * 2) = element__k(i_element, 0, 1) + dl__k(2 * i_element + 0, 1 + i_element * 2)
    element__k(i_element, 0, 2) = -constante * 12: dl__k(2 * i_element + 0, 2 + i_element * 2) = element__k(i_element, 0, 2) + dl__k(2 * i_element + 0, 2 + i_element * 2)
    element__k(i_element, 0, 3) = constante * element__long(i_element) * 6: dl__k(2 * i_element + 0, 3 + i_element * 2) = element__k(i_element, 0, 3) + dl__k(2 * i_element + 0, 3 + i_element * 2)
    element__k(i_element, 1, 0) = constante * element__long(i_element) * 6: dl__k(2 * i_element + 1, 0 + i_element * 2) = element__k(i_element, 1, 0) + dl__k(2 * i_element + 1, 0 + i_element * 2)
    element__k(i_element, 1, 1) = constante * element__long(i_element) ^ 2 * 4: dl__k(2 * i_element + 1, 1 + i_element * 2) = element__k(i_element, 1, 1) + dl__k(2 * i_element + 1, 1 + i_element * 2)
    element__k(i_element, 1, 2) = -constante * element__long(i_element) * 6: dl__k(2 * i_element + 1, 2 + i_element * 2) = element__k(i_element, 1, 2) + dl__k(2 * i_element + 1, 2 + i_element * 2)
    element__k(i_element, 1, 3) = constante * element__long(i_element) ^ 2 * 2: dl__k(2 * i_element + 1, 3 + i_element * 2) = element__k(i_element, 1, 3) + dl__k(2 * i_element + 1, 3 + i_element * 2)
    element__k(i_element, 2, 0) = -constante * 12: dl__k(2 * i_element + 2, 0 + i_element * 2) = element__k(i_element, 2, 0) + dl__k(2 * i_element + 2, 0 + i_element * 2)
    element__k(i_element, 2, 1) = -constante * element__long(i_element) * 6: dl__k(2 * i_element + 2, 1 + i_element * 2) = element__k(i_element, 2, 1) + dl__k(2 * i_element + 2, 1 + i_element * 2)
    element__k(i_element, 2, 2) = constante * 12: dl__k(2 * i_element + 2, 2 + i_element * 2) = element__k(i_element, 2, 2) + dl__k(2 * i_element + 2, 2 + i_element * 2)
    element__k(i_element, 2, 3) = -constante * element__long(i_element) * 6: dl__k(2 * i_element + 2, 3 + i_element * 2) = element__k(i_element, 2, 3) + dl__k(2 * i_element + 2, 3 + i_element * 2)
    element__k(i_element, 3, 0) = constante * element__long(i_element) * 6: dl__k(2 * i_element + 3, 0 + i_element * 2) = element__k(i_element, 3, 0) + dl__k(2 * i_element + 3, 0 + i_element * 2)
    element__k(i_element, 3, 1) = constante * element__long(i_element) ^ 2 * 2: dl__k(2 * i_element + 3, 1 + i_element * 2) = element__k(i_element, 3, 1) + dl__k(2 * i_element + 3, 1 + i_element * 2)
    element__k(i_element, 3, 2) = -constante * element__long(i_element) * 6: dl__k(2 * i_element + 3, 2 + i_element * 2) = element__k(i_element, 3, 2) + dl__k(2 * i_element + 3, 2 + i_element * 2)
    element__k(i_element, 3, 3) = constante * element__long(i_element) ^ 2 * 4: dl__k(2 * i_element + 3, 3 + i_element * 2) = element__k(i_element, 3, 3) + dl__k(2 * i_element + 3, 3 + i_element * 2)
Next i_element

'Afficher la matrice de rigidité
Debug.Print "2. Matrice de rigidité globale par élément"
For i_element = 0 To nb_element - 1
    Debug.Print "Element " & i_element
    For i = 0 To 3
        Debug.Print dl__k(i_element * 2 + i, i_element * 2 + 0), dl__k(i_element * 2 + i, i_element * 2 + 1), dl__k(i_element * 2 + i, i_element * 2 + 2), dl__k(i_element * 2 + i, i_element * 2 + 3)
    Next i
Next i_element

'3 - VECTEUR FORCE GLOBAL

ReDim dl__f(nb_dl - 1)
ReDim element__f(nb_element - 1, 4)
For i_noeud = 0 To nb_noeud - 1 'charge ponctuelle
    dl__f(i_noeud * 2) = dl__f(i_noeud * 2) + noeud__fext(i_noeud) 'dl uy
    dl__f(i_noeud * 2 + 1) = 0 'dl rotz
Next i_noeud
    
For i_element = 0 To nb_element - 1 'charge linéaire
    dle__feq(0) = -(element__qext(i_element) * element__long(i_element) / 2)
    dle__feq(1) = -(element__qext(i_element) * element__long(i_element) ^ 2 / 12)
    dle__feq(2) = -(element__qext(i_element) * element__long(i_element) / 2)
    dle__feq(3) = element__qext(i_element) * element__long(i_element) ^ 2 / 12
    For i_dle = 0 To 3
        dl__f(i_element * 2 + i_dle) = dl__f(i_element * 2 + i_dle) - dle__feq(i_dle)
        element__f(i_element, i_dle) = element__f(i_element, i_dle) + dle__feq(i_dle)
    Next i_dle
Next i_element

''Afficher le vecteur force global et forces par element
'Debug.Print "3.1 - Vecteur force global (dl_f)"
'For i_noeud = 0 To nb_noeud - 1
'    Debug.Print "noeud " & i_noeud & " - dl_uy : " & dl__f(i_noeud * 2)
'    Debug.Print "noeud " & i_noeud & " - dl_rotz : " & dl__f(i_noeud * 2 + 1)
'Next i_noeud
'Debug.Print "3.2 - Vecteur force element (element__f)"
'For i_element = 0 To nb_element - 1
'    Debug.Print "element " & i_element & " - forces :", element__f(i_element, 0), element__f(i_element, 1), element__f(i_element, 2), element__f(i_element, 3)
'Next i_element

'4 - CONDITIONS LIMITES

For i_noeud = 0 To nb_noeud - 1
    If noeud__appui(i_noeud) Then 'simplification de la matrice de rigidité globale (dl__k) et du vecteur force global (dl_f)
        For i_dl = 0 To nb_dl - 1
            dl__k(i_dl, i_noeud * 2) = 0
            dl__k(i_noeud * 2, i_dl) = 0
        Next i_dl
        dl__k(i_noeud * 2, i_noeud * 2) = 1 'dl_uy bloqué pour une liaison en appui
        dl__f(i_noeud * 2) = 0 'dl_uy bloqué pour une liaison en appui
    End If
Next i_noeud

''Afficher la matrice de rigidité globale et le vecteur force global
'Debug.Print "Matrice de rigidité globale par élément"
'For i_element = 0 To nb_element - 1
'    Debug.Print "Element " & i_element
'    For i = 0 To 3
'        Debug.Print dl__k(i_element * 2 + i, i_element * 2 + 0), dl__k(i_element * 2 + i, i_element * 2 + 1), dl__k(i_element * 2 + i, i_element * 2 + 2), dl__k(i_element * 2 + i, i_element * 2 + 3)
'    Next i
'Next i_element
'Debug.Print "Vecteur force global (dl_f)"
'For i_noeud = 0 To nb_noeud - 1
'    Debug.Print "noeud " & i_noeud & " - dl_uy : " & dl__f(i_noeud * 2)
'    Debug.Print "noeud " & i_noeud & " - dl_rotz : " & dl__f(i_noeud * 2 + 1)
'Next i_noeud


'5 - DEPLACEMENTS

ReDim dl__u(nb_dl - 1)
dl__u = dl__f

'élimination
For i_dl = 0 To nb_dl - 2 'N
    For ii_dl = i_dl + 1 To nb_dl - 1 'N+1
        ratio = dl__k(i_dl, ii_dl) / dl__k(i_dl, i_dl)
        dl__u(ii_dl) = dl__u(ii_dl) - ratio * dl__u(i_dl)
        For jj_dl = ii_dl To nb_dl - 1
            dl__k(ii_dl, jj_dl) = dl__k(ii_dl, jj_dl) - ratio * dl__k(i_dl, jj_dl)
        Next jj_dl
    Next ii_dl
Next i_dl

'substitution arrière
dl__u(nb_dl - 1) = dl__u(nb_dl - 1) / dl__k(nb_dl - 1, nb_dl - 1)
For i_dl = nb_dl - 2 To 0 Step -1
    diminuteur = 0
    For jj_dl = i_dl + 1 To nb_dl - 1
        diminuteur = diminuteur + dl__k(i_dl, jj_dl) * dl__u(jj_dl)
    Next jj_dl
    dl__u(i_dl) = (dl__u(i_dl) - diminuteur) / dl__k(i_dl, i_dl)
Next i_dl

''Afficher le vecteur déplacement global (dl__u)
'Debug.Print "Vecteur déplacement global (dl_u)"
'For i_noeud = 0 To nb_noeud - 1
'    Debug.Print "noeud " & i_noeud & " - dl_uy : " & dl__u(i_noeud * 2)
'    Debug.Print "noeud " & i_noeud & " - dl_rotz : " & dl__u(i_noeud * 2 + 1)
'    Debug.Print ""
'Next i_noeud


'6 - EFFORTS ELEMENTAIRES

For i_element = 0 To nb_element - 1
    For i_dle = 0 To 3
        For j_dle = 0 To 3
            element__f(i_element, i_dle) = element__f(i_element, i_dle) + element__k(i_element, i_dle, j_dle) * dl__u(i_element * 2 + j_dle)
        Next j_dle
    Next i_dle
Next i_element

''Afficher les vecteurs de force élémentaires
'Debug.Print "Vecteurs des forces élémentaires (element__f)"
'For i_element = 0 To nb_element - 1
'    Debug.Print "Element " & i_element & " FY " & element__f(i_element, 0), "MZ " & element__f(i_element, 1)
'    Debug.Print "Element " & i_element & " FY " & element__f(i_element, 2), "MZ " & element__f(i_element, 3)
'    Debug.Print ""
'Next i_element


'7 - REACTION LIAISON

ReDim noeud__ry(nb_noeud - 1)
ReDim noeud__mz(nb_noeud - 1)
For i_noeud = 0 To nb_noeud - 1
    If noeud__appui(i_noeud) Then
        nb_appui = nb_appui + 1
        If i_noeud < nb_element Then 'ORIGINE verif porte à faux faitage
            noeud__ry(i_noeud) = element__f(i_noeud, 0)
            'noeud__mz(i_noeud) = element__f(i_noeud, 1) pas encastrement
        End If
        If i_noeud > 0 Then 'EXTREMITE verif porte à faux égout
            noeud__ry(i_noeud) = element__f(i_noeud - 1, 2) + noeud__ry(i_noeud)
            'noeud__mz(i_noeud) = element__f(i_noeud - 1, 3) + noeud__mz(i_noeud) pas encastrement
        End If
    End If
Next i_noeud

''Afficher les réaction aux liaison
'For i_noeud = 0 To nb_noeud - 1
'    If noeud__appui(i_noeud) Then
'        Debug.Print "noeud " & i_noeud & " - Ry : " & noeud__ry(i_noeud)
'        'Debug.Print "noeud " & i_noeud & " - Mz : " & noeud__mz(i_noeud)
'    End If
'Next i_noeud


'8 - FLECHE ET MOMENT MAX

nb_travee = nb_appui - 1
ReDim travee__uy_max(nb_travee - 1, 1)
ReDim travee__mfz_0(nb_travee - 1, 1)
ReDim travee__mfz_pos_max(nb_travee - 1, 1)
ReDim travee__mfz_neg_min(nb_travee - 1, 1)
While noeud__appui(0 + paf_egout) = False: paf_egout = paf_egout + 1: Wend
While noeud__appui(nb_noeud - 1 - paf_faitage) = False: paf_faitage = paf_faitage + 1: Wend
For i_travee = 0 To nb_travee - 1
    travee__mfz_0(i_travee, 0) = 1E+99 'cherche inflexion min abso
    travee__mfz_pos_max(i_travee, 0) = -1E+99 'cherche max
    travee__mfz_neg_min(i_travee, 0) = 1E+99 'cherche min
Next i_travee
i_travee = 0
For i_element = 0 + paf_egout To nb_element - 1 - paf_faitage
    'Debug.Print i_element, i_travee
    longueur = element__long(i_element)
    young = element__young(i_element)
    iz = element__iz(i_element)
    qy_ext = element__qext(i_element)
    f2 = element__f(i_element, 2)
    f3 = element__f(i_element, 3)
    uy_origine = dl__u(i_element * 2)
    rotz_origine = dl__u(i_element * 2 + 1)
    uy_extremite = dl__u((i_element + 1) * 2)
    rotz_extremite = dl__u((i_element + 1) * 2 + 1)
    nb_x = element__long(i_element) * 1000 'm > mm
    ReDim x__uy(nb_x)
    ReDim x__ty(nb_x)
    ReDim x__mfz(nb_x)
    For x = 0 To nb_x
        xi = x / nb_x
        x__uy(x) = uy_origine * (1 - 3 * xi ^ 2 + 2 * xi ^ 3) _
                 + rotz_origine * (xi - 2 * xi ^ 2 + xi ^ 3) * longueur _
                 + uy_extremite * (3 * xi ^ 2 - 2 * xi ^ 3) _
                 + rotz_extremite * (-xi ^ 2 + xi ^ 3) * longueur
        'x__ty(x) = f2
        x__mfz(x) = f3 + longueur * f2 * (1 - xi)
        If Abs(qy_ext) > 0.1 Then
            x__uy(x) = x__uy(x) + 0.5 * (qy_ext * longueur ^ 2 * longueur / (12 * young * iz)) * longueur * xi ^ 2 * (1 - xi) ^ 2
            'x__ty(x) = x__ty(x) + qy_ext * longueur * (1 - xi)
            x__mfz(x) = x__mfz(x) + 0.5 * qy_ext * longueur ^ 2 * (1 - xi) ^ 2
        End If
        abscisse = longueur * xi + noeud__axe(i_element)
        If Round(Abs(x__uy(x)), 6) > Round(Abs(travee__uy_max(i_travee, 0)), 6) Then travee__uy_max(i_travee, 0) = x__uy(x): travee__uy_max(i_travee, 1) = abscisse
        If Round(Abs(x__mfz(x)), 6) < Round(Abs(travee__mfz_0(i_travee, 0)), 6) Then travee__mfz_0(i_travee, 0) = x__mfz(x): travee__mfz_0(i_travee, 1) = abscisse
        If Round(x__mfz(x), 6) > Round(travee__mfz_pos_max(i_travee, 0), 6) Then: travee__mfz_pos_max(i_travee, 0) = x__mfz(x): travee__mfz_pos_max(i_travee, 1) = abscisse
        If Round(x__mfz(x), 6) < Round(travee__mfz_neg_min(i_travee, 0), 6) Then: travee__mfz_neg_min(i_travee, 0) = x__mfz(x): travee__mfz_neg_min(i_travee, 1) = abscisse
    Next x
    If noeud__appui(i_element + 1) And i_travee < nb_travee - 1 Then i_travee = i_travee + 1
Next i_element

'9 - RESULTATS

Set output = ThisWorkbook.Sheets("output"): output.Cells.ClearContents: i = 2
Set output_concat = ThisWorkbook.Sheets("output").Range("A1")
delta_row = 16

output.Range("A" & 1 + delta_row).Value = "End deflection [m]"
output.Range("B" & 1 + delta_row).Value = "Max span deflection [m]"
output.Range("C" & 1 + delta_row).Value = "x [m]"
output.Range("A" & 2 + delta_row).Value = dl__u(0): output.Range("A" & 3 + delta_row).Value = dl__u((nb_noeud - 1) * 2)
output_0 = output_0 & ":" & dl__u(0): output_0 = output_0 & ":" & dl__u((nb_noeud - 1) * 2)
For i_travee = LBound(travee__uy_max) To UBound(travee__uy_max)
    output.Range("B" & i + i_travee + delta_row).Value = Round(travee__uy_max(i_travee, 0), 6)
    output.Range("C" & i + i_travee + delta_row).Value = Round(travee__uy_max(i_travee, 1), 4)
    output_1 = output_1 & ":" & Round(travee__uy_max(i_travee, 0), 6)
    output_2 = output_2 & ":" & Round(travee__uy_max(i_travee, 1), 4)
Next i_travee

output.Range("D" & 1 + delta_row).Value = "Fy (support) [N]"
output.Range("E" & 1 + delta_row).Value = "Mfz (support) [N.m]"

i_travee = 0: For i_noeud = 0 To nb_noeud - 1
    If noeud__appui(i_noeud) Then
        output.Range("D" & i + i_travee + delta_row).Value = Round(noeud__ry(i_noeud), 3)
        output.Range("E" & i + i_travee + delta_row).Value = Round(element__f(i_noeud, 1), 3)
        output_3 = output_3 & ":" & Round(noeud__ry(i_noeud), 3)
        output_4 = output_4 & ":" & Round(element__f(i_noeud, 1), 3)
        i_travee = i_travee + 1
    End If
Next i_noeud

output.Range("F" & 1 + delta_row).Value = "Mfz |min| span [N.m]"
output.Range("G" & 1 + delta_row).Value = "x [m]"
output.Range("H" & 1 + delta_row).Value = "Mfz max span [N.m]"
output.Range("I" & 1 + delta_row).Value = "x [m]"
output.Range("J" & 1 + delta_row).Value = "Mfz min span [N.m]"
output.Range("K" & 1 + delta_row).Value = "x [m]"

i_travee = 0
For i_travee = LBound(travee__mfz_0) To UBound(travee__mfz_0)
    output.Range("F" & i + i_travee + delta_row).Value = Round(travee__mfz_0(i_travee, 0), 6)
    output.Range("G" & i + i_travee + delta_row).Value = Round(travee__mfz_0(i_travee, 1), 4)
    output.Range("H" & i + i_travee + delta_row).Value = Round(travee__mfz_pos_max(i_travee, 0), 6)
    output.Range("I" & i + i_travee + delta_row).Value = Round(travee__mfz_pos_max(i_travee, 1), 4)
    output.Range("J" & i + i_travee + delta_row).Value = Round(travee__mfz_neg_min(i_travee, 0), 6)
    output.Range("K" & i + i_travee + delta_row).Value = Round(travee__mfz_neg_min(i_travee, 1), 4)
    
    output_5 = output_5 & ":" & Round(travee__mfz_0(i_travee, 0), 6)
    output_6 = output_6 & ":" & Round(travee__mfz_0(i_travee, 1), 4)
    output_7 = output_7 & ":" & Round(travee__mfz_pos_max(i_travee, 0), 6)
    output_8 = output_8 & ":" & Round(travee__mfz_neg_min(i_travee, 0), 6)
Next i_travee

'output.Columns("A:Z").AutoFit
output_concat.Value = Mid(output_0, 2) & ";" & Mid(output_1, 2) & ";" & Mid(output_2, 2) & ";" & Mid(output_3, 2) & ";" & Mid(output_4, 2) & ";" & Mid(output_5, 2) & ";" & Mid(output_6, 2) & ";" & Mid(output_7, 2) & ";" & Mid(output_8, 2)

'End
'Debug.Print "Travée", "Flèche (m)", "Abscisse (m)"
'For i_travee = LBound(travee__uy_max) To UBound(travee__uy_max)
'    Debug.Print "Travée " & i_travee + 1, Round(travee__uy_max(i_travee, 0), 6), Round(travee__uy_max(i_travee, 1), 4)
'Next i_travee
'Debug.Print "Flèche (m) égout : ", Round(dl__u(0), 6)
'Debug.Print "Flèche (m) faitage : ", Round(dl__u(nb_dl - 1), 6)
'
'For i_noeud = 0 To nb_noeud - 1
'    If noeud__appui(i_noeud) Then
'        Debug.Print "Ry (N) noeud " & i_noeud & " : ", Round(noeud__ry(i_noeud), 3)
'    End If
'Next i_noeud
'
'For i_element = 0 To nb_element - 1
'    If noeud__appui(i_element) Then
'        Debug.Print "Mfz (N.m) noeud " & i_element & " : ", Round(element__f(i_element, 1), 3)
'    End If
'Next i_element
'Debug.Print "Travée", "Moment (N.m)", "Abscisse (m)"
'For i_travee = LBound(travee__mfz_0) To UBound(travee__mfz_0)
'    Debug.Print "Travée " & i_travee + 1, Round(travee__mfz_0(i_travee, 0), 6), Round(travee__mfz_0(i_travee, 1), 4)
'Next i_travee

'End

'10 - BONUS GRAPHES
'Debug.Print "9 - GRAPHE"
Set graph = ThisWorkbook.Sheets("output_graph")
'graph.Cells.Clear: graph.Range("A1").Value = "x (mm)": graph.Range("B1").Value = "uy (m)": graph.Range("C1").Value = "Ty (N))": graph.Range("D1").Value = "Mfz (N.m)": graph.Range("E1").Value = "Liaison (appui)"
concat_graph0 = "x (mm)"
concat_graph1 = "uy (m)"
concat_graph2 = "Ty (N)"
concat_graph3 = "Mfz (N.m)"
concat_graph4 = "Liaison (appui)"
i = 2
resolution = extremite_poutres(UBound(extremite_poutres)) / 2997 'nb lignes
For i_element = 0 To nb_element - 1
    longueur = element__long(i_element)
    young = element__young(i_element)
    iz = element__iz(i_element)
    qy_ext = element__qext(i_element)
    f2 = element__f(i_element, 2)
    f3 = element__f(i_element, 3)
    uy_origine = dl__u(i_element * 2)
    rotz_origine = dl__u(i_element * 2 + 1)
    uy_extremite = dl__u((i_element + 1) * 2)
    rotz_extremite = dl__u((i_element + 1) * 2 + 1)
    nb_x = Round(element__long(i_element) / resolution + 0.49999, 1)
    ReDim x__uy(nb_x)
    ReDim x__ty(nb_x)
    ReDim x__mfz(nb_x)
    'If noeud__appui(i_element) Then graph.Cells(i, 5).Value = i_element / 1000000
    If noeud__appui(i_element) Then concat_graph_appui = i_element / 100000 Else concat_graph_appui = ""
    For x = 0 To nb_x
        xi = x / nb_x
        x__uy(x) = uy_origine * (1 - 3 * xi ^ 2 + 2 * xi ^ 3) _
                 + rotz_origine * (xi - 2 * xi ^ 2 + xi ^ 3) * longueur _
                 + uy_extremite * (3 * xi ^ 2 - 2 * xi ^ 3) _
                 + rotz_extremite * (-xi ^ 2 + xi ^ 3) * longueur
        x__ty(x) = f2
        x__mfz(x) = f3 + longueur * f2 * (1 - xi)
        If Abs(qy_ext) > 0.1 Then
            x__uy(x) = x__uy(x) + 0.5 * (qy_ext * longueur ^ 2 * longueur / (12 * young * iz)) * longueur * xi ^ 2 * (1 - xi) ^ 2
            x__ty(x) = x__ty(x) + qy_ext * longueur * (1 - xi)
            x__mfz(x) = x__mfz(x) + 0.5 * qy_ext * longueur ^ 2 * (1 - xi) ^ 2
        End If
        'graph.Cells(i, 1).Value = (i - 2) * resolution * 1000 'mm
        'graph.Cells(i, 2).Value = x__uy(x)
        'graph.Cells(i, 3).Value = x__ty(x)
        'graph.Cells(i, 4).Value = x__mfz(x)
        If x > 0 Or i_element = 0 Then concat_graph0 = concat_graph0 & ":" & Round((longueur * xi + noeud__axe(i_element)) * 1000, 1) 'mm
        If x > 0 Or i_element = 0 Then concat_graph1 = concat_graph1 & ":" & Round(x__uy(x), 4)
        If x > 0 Or i_element = 0 Then concat_graph2 = concat_graph2 & ":" & Round(x__ty(x), 1)
        If x > 0 Or i_element = 0 Then concat_graph3 = concat_graph3 & ":" & Round(x__mfz(x), 2)
        If (x > 0 Or i_element = 0) And x < nb_x Then concat_graph4 = concat_graph4 & ":"
        i = i + 1
    Next x
    i = i - 1
    'If noeud__appui(i_element + 1) Then graph.Cells(i, 5).Value = i_element + 1
    If noeud__appui(i_element + 1) Then concat_graph_appui = i_element / 100000 Else concat_graph_appui = ""
    concat_graph4 = concat_graph4 & ":" & concat_graph_appui
Next i_element

graph.Range("A1").Value = concat_graph0
graph.Range("B1").Value = concat_graph1
graph.Range("C1").Value = concat_graph2
graph.Range("D1").Value = concat_graph3
graph.Range("E1").Value = concat_graph4

End Function



