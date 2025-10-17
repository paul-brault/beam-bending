Attribute VB_Name = "functions"
Function compliance(data As Variant)
'pas besoin de vérifier car algorithme accepte :
    'abscisse non croissant : force ponctuelle et/ou lineaire
    'chevauchement : force ponctuelle et/ou lineaire
compliance = "ok"

If UBound(Split(data, ";")) <> 8 Then compliance = "Erreur : Nombre de séparateurs ';' incorrecte": Exit Function

data_split = Split(data, ";") '0axe_appui;1extremite_poutre;2young_poutre;3iz_poutre;4axe_ponctuelle;5force_ponctuelle;6origine_linéaire;7extremite_linéaire;8force_linéaire
noeud__axe = Split("0:" & data_split(0) & ":" & data_split(1) & ":" & data_split(4) & ":" & data_split(6) & ":" & data_split(7), ":")
axe_appuis = Split(data_split(0), ":")
extremite_poutre = Split(data_split(1), ":")
young_poutre = Split(data_split(2), ":")
iz_poutre = Split(data_split(3), ":")
axe_ponctuelle = Split(data_split(4), ":")
force_ponctuelle = Split(data_split(5), ":")
origine_lineaire = Split(data_split(6), ":")
extremite_lineaire = Split(data_split(7), ":")
force_lineaire = Split(data_split(8), ":")

For i = LBound(extremite_poutre) + 1 To UBound(extremite_poutre)
    If Round(extremite_poutre(i - 1), 5) > Round(extremite_poutre(i), 5) Then compliance = "Erreur : Extremité poutre non croissant": Exit Function
Next i

If UBound(extremite_poutre) <> UBound(young_poutre) Or UBound(extremite_poutre) <> UBound(iz_poutre) Then compliance = "Erreur : Nombre extremité, young ou iz poutre incohérent": Exit Function

If UBound(force_lineaire) <> UBound(origine_lineaire) Or UBound(force_lineaire) <> UBound(extremite_lineaire) Then compliance = "Erreur : Nombre origine, extrémité ou force linéaire incohérent": Exit Function

For i = 0 To UBound(axe_appuis)
    If Round(axe_appuis(i), 5) > Round(extremite_poutre(UBound(extremite_poutre)), 5) Then compliance = "Erreur : Axe appuis > extremité poutre": Exit Function
Next i

For i = 0 To UBound(axe_ponctuelle)
    If Round(axe_ponctuelle(i), 5) > Round(extremite_poutre(UBound(extremite_poutre)), 5) Then compliance = "Erreur : Axe ponctuelle > extremité poutre": Exit Function
Next i

For i = 0 To UBound(origine_lineaire)
    If Round(origine_lineaire(i), 5) > Round(extremite_poutre(UBound(extremite_poutre)), 5) Then compliance = "Erreur : Origine linéaire > extremité poutre": Exit Function
Next i

For i = 0 To UBound(extremite_lineaire)
    If Round(extremite_lineaire(i), 5) > Round(extremite_poutre(UBound(extremite_poutre)), 5) Then compliance = "Erreur : Extrémité linéaire > extremité poutre": Exit Function
Next i

If UBound(axe_appuis) < 1 Then compliance = "Erreur : Nombre appuis < 2": Exit Function

For i = 0 To UBound(force_ponctuelle)
    fy = fy + force_ponctuelle(i)
Next i
For i = 0 To UBound(force_lineaire)
    fy = fy + force_lineaire(i)
Next i
If fy = 0 Then compliance = "Erreur : Aucun chargement": Exit Function

If UBound(noeud__axe) > 100 Then compliance = "Erreur : Nombre de noeuds > 100": Exit Function

End Function

Public Sub ArraySort(vArray As Variant, inLow As Long, inHi As Long)
    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long

    tmpLow = inLow
    tmpHi = inHi

    pivot = vArray((inLow + inHi) \ 2)

    While (tmpLow <= tmpHi)
        While (Round(vArray(tmpLow), 10) < Round(pivot, 10) And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend

        While (Round(pivot, 10) < Round(vArray(tmpHi), 10) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend

        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend

    If (inLow < tmpHi) Then ArraySort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then ArraySort vArray, tmpLow, inHi
End Sub

Public Sub ArrayUnique(ByRef vArray As Variant)
    Dim uniqueList() As Variant
    Dim uniqueCount As Long
    Dim i As Long

    If UBound(vArray) < LBound(vArray) Then Exit Sub ' La liste est vide
    
    ' Initialiser la liste des éléments uniques
    ReDim uniqueList(0 To UBound(vArray))
    uniqueCount = 0

    ' Ajouter le premier élément à la liste des éléments uniques
    uniqueList(uniqueCount) = vArray(LBound(vArray))
    uniqueCount = uniqueCount + 1

    ' Parcourir la liste pour ajouter les éléments uniques
    For i = LBound(vArray) + 1 To UBound(vArray)
        If Round(CDbl(vArray(i)), 5) <> Round(CDbl(vArray(i - 1)), 5) Then
        'If vArray(i) <> vArray(i - 1) Then
            ReDim Preserve uniqueList(0 To uniqueCount)
            uniqueList(uniqueCount) = vArray(i)
            uniqueCount = uniqueCount + 1
        End If
    Next i

    ' Réduire la taille de la liste originale
    ReDim vArray(0 To uniqueCount - 1)
    
    ' Copier les éléments uniques dans la liste originale
    For i = LBound(uniqueList) To UBound(uniqueList)
        vArray(i) = uniqueList(i) 'éviter les erreur virgule flotante VBA
    Next i
End Sub


'Dim noeud__dl_id()
'Dim element__noeud_id()

'ReDim noeud__dl_id(nb_noeud - 1, 1)
'For i = 0 To nb_noeud - 1
'    For j = 0 To 1 'uy and rotz
'        noeud__dl_id(i, j) = i * 2 + j
'        'Debug.Print noeud__dl_id(i, j)
'    Next j
'Next i
'
'nb_element = nb_noeud - 1
'ReDim element__noeud_id(nb_element - 1, 1)
'For i = 0 To nb_element - 1
'    For j = 0 To 1 'origine and extremite
'        element__noeud_id(i, j) = i + j
'        'Debug.Print element__noeud_id(i, j)
'    Next j
'Next i

Public Function NormalizeDataLocale(ByVal data As String) As String
    Dim decSep As String, otherSep As String
    decSep = Application.International(xlDecimalSeparator)
    otherSep = IIf(decSep = ".", ",", ".")

    ' supprime espaces et séparateurs de milliers classiques
    data = Replace(data, " ", "")
    data = Replace(data, otherSep, "")

    ' convertit l'autre symbole décimal vers la locale
    If decSep = "," Then
        data = Replace(data, ".", ",")
    Else
        data = Replace(data, ",", ".")
    End If

    NormalizeDataLocale = data
End Function

Private Function ParseDbl(ByVal s As String) As Double
    Dim decSep As String, thouSep As String
    decSep = Application.International(xlDecimalSeparator)
    thouSep = IIf(decSep = ".", ",", ".")
    s = Trim$(s)
    s = Replace(s, thouSep, "")                ' supprime séparateur de milliers
    s = Replace(s, IIf(decSep = ".", ",", "."), decSep) ' convertit l'autre symbole vers la locale
    If Left$(s, 1) = decSep Then s = "0" & s
    If Right$(s, 1) = decSep Then s = s & "0"
    If Not IsNumeric(s) Then Err.Raise vbObjectError + 513, , "Valeur non numérique: '" & s & "'"
    ParseDbl = CDbl(s)
End Function


