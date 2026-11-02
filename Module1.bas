Option Explicit

' Adresse du serveur OpenFoodFacts pour chercher des produits
Const URL_PRODUITS As String = "https://world.openfoodfacts.org/cgi/search.pl"

' Adresse du serveur OpenPrices pour récupérer les prix
Const URL_PRIX As String = "https://prices.openfoodfacts.org/api/v1/prices"

' On récupère 100 produits par page
Const NB_PAR_PAGE As Integer = 100

' On s'arrête après 5 pages maximum (soit 500 produits max)
Const NB_PAGES_MAX As Integer = 5


' =============================================
' MACRO PRINCIPALE : RECHERCHER DES PRODUITS
' =============================================
Sub RechercherProduits()

    ' ---- Etape 1 : on lit ce que l'utilisateur a tapé ----

    ' On va chercher la feuille "Recherche" où l'utilisateur saisit son mot
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Recherche")

    ' On lit le mot tapé en D5 et on retire les espaces inutiles avant/après
    Dim terme As String
    terme = Trim(ws.Range("D5").Value)

    ' Si l'utilisateur n'a rien tapé, on l'avertit et on arrête tout
    If terme = "" Then
        MsgBox "Saisissez un mot à rechercher en D5", vbExclamation
        Exit Sub
    End If


    ' ---- Etape 2 : on prépare la feuille qui va recevoir les résultats ----

    ' On essaie de récupérer la feuille "Resultats" si elle existe déjà
    Dim wsR As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Sheets("Resultats")
    On Error GoTo 0

    ' Si elle n'existe pas encore, on la crée automatiquement
    If wsR Is Nothing Then
        Set wsR = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsR.Name = "Resultats"
    End If

    ' On efface les anciens résultats et on remet les titres de colonnes
    wsR.Cells.Clear
    EcrireEntetes wsR


    ' ---- Etape 3 : on traduit le mot dans plusieurs langues ----
    ' OpenFoodFacts contient des produits du monde entier
    ' En cherchant "chocolate", "schokolade", "cioccolato"... on trouve bien plus de produits
    ' qu'en cherchant uniquement "chocolat" en français

    Statut ws, "Traduction du mot en plusieurs langues..."
    Dim termes() As String
    termes = TraduireMultiLangues(terme)


    ' ---- Etape 4 : on lance la recherche pour chaque traduction ----

    ' Cette variable mémorise les codes-barres déjà affichés pour éviter les doublons
    Dim codesVus As String
    codesVus = "|"

    ' On commence à écrire les résultats à la ligne 3 (la ligne 2 contient les titres)
    Dim ligne As Long
    ligne = 3

    ' Déclaration de toutes les variables utilisées dans la boucle
    Dim t            As Integer  ' Numéro du terme en cours (0 = français, 1 = anglais, etc.)
    Dim motEnCours   As String   ' Le mot dans la langue en cours
    Dim urlBase      As String   ' L'URL de base pour cette recherche
    Dim urlRecherche As String   ' L'URL complète avec le numéro de page
    Dim reponse      As String   ' La réponse brute reçue du serveur (format JSON)
    Dim produits()   As String   ' Liste des produits extraits de la réponse
    Dim i            As Integer  ' Numéro du produit en cours
    Dim p            As String   ' Les données d'un produit (en JSON)
    Dim code         As String   ' Le code-barres du produit
    Dim valide       As Boolean  ' Est-ce que ce produit correspond à notre recherche ?
    Dim page         As Integer  ' Numéro de la page en cours
    Dim nbPages      As Integer  ' Nombre total de pages disponibles
    Dim pageCount    As Long     ' Valeur lue dans la réponse JSON

    ' On boucle sur chaque traduction du mot
    For t = 0 To UBound(termes)

        motEnCours = termes(t)

        ' On construit l'URL de recherche avec le mot traduit
        ' search_terms= indique à OpenFoodFacts ce qu'on cherche
        urlBase = URL_PRODUITS & "?search_terms=" & EncodeURL(motEnCours)

        ' On ajoute les paramètres communs :
        ' action=process : lance la recherche
        ' json=1 : on veut la réponse en JSON (format lisible par le code)
        ' page_size : combien de produits par page
        ' fields : on ne demande que les infos dont on a besoin (pour aller plus vite)
        urlBase = urlBase & "&action=process&json=1&page_size=" & NB_PAR_PAGE & _
                  "&fields=product_name,brands,quantity,countries,nutriscore_grade,ecoscore_grade,code,image_url"

        ' On commence à la première page et on avancera page par page
        page = 1
        nbPages = 1

        ' Boucle de pagination : on continue tant qu'il reste des pages à charger
        Do While page <= nbPages And page <= NB_PAGES_MAX

            ' On affiche la progression dans la cellule D13
            Statut ws, "Recherche : """ & motEnCours & """" & _
                       " - Page " & page & "/" & nbPages & _
                       " (langue " & t + 1 & "/" & UBound(termes) + 1 & ")"

            ' On ajoute le numéro de page à l'URL
            urlRecherche = urlBase & "&page=" & page

            ' On envoie la requête au serveur et on récupère la réponse
            reponse = GetHTTP(urlRecherche)

            ' On attend 800ms avant la prochaine requête
            ' pour ne pas surcharger les serveurs et eviter d'être bloqué
            Wait 800

            ' Si le serveur a bien répondu (réponse non vide)
            If reponse <> "" Then

                ' Dès la première page, on lit le nombre total de pages disponibles
                ' pour savoir combien de fois il faudra boucler
                If page = 1 Then
                    Dim pageCountStr As String
                    pageCountStr = Valeur(reponse, "page_count")
                    If pageCountStr <> "" Then
                        pageCount = CLng(pageCountStr)
                        ' On ne dépasse pas le maximum qu'on s'est fixé
                        nbPages = CInt(WorksheetFunction.Min(pageCount, NB_PAGES_MAX))
                    End If
                End If

                ' On extrait la liste des produits du JSON
                ' puis on la découpe en produits individuels
                produits = SplitObjets(ExtraireTableau(reponse, "products"))

                ' On examine chaque produit un par un
                For i = 0 To UBound(produits)

                    p = produits(i)

                    ' On lit le code-barres du produit
                    ' (on utilise ValeurTexte car les codes-barres sont de très grands nombres
                    '  que VBA risquerait de tronquer s'il les traitait comme des chiffres)
                    code = ValeurTexte(p, "code")

                    ' On vérifie que le nom du produit contient bien le mot recherché
                    ' (l'API peut parfois renvoyer des produits hors sujet)
                    valide = False
                    If InStr(LCase(CleanTxt(Valeur(p, "product_name"))), LCase(motEnCours)) > 0 Then
                        valide = True
                    End If

                    ' On ajoute le produit dans la feuille si :
                    ' - il correspond à notre recherche (valide = True)
                    ' - il a un code-barres (sinon on ne peut pas le identifier)
                    ' - on ne l'a pas déjà affiché (anti-doublon)
                    If valide And code <> "" And InStr(codesVus, "|" & code & "|") = 0 Then
                        codesVus = codesVus & code & "|"
                        AjouterLigne wsR, ligne, p, code
                        ligne = ligne + 1
                    End If

                Next i

            End If

            ' On passe à la page suivante
            page = page + 1

        Loop

    Next t


    ' ---- Etape 5 : mise en forme et chargement des prix ----

    ' On ajuste les largeurs de colonnes pour une belle présentation
    MiseEnForme wsR

    ' On bascule sur la feuille résultats pour que l'utilisateur la voit
    wsR.Activate

    ' Si on a trouvé au moins un produit, on charge les prix depuis OpenPrices
    ' Sinon on affiche "Aucun résultat"
    If ligne > 3 Then
        ChargePrix wsR, ligne - 1, ws
    Else
        Statut ws, "Aucun résultat trouvé pour : " & terme
    End If

End Sub


' =============================================
' MACRO 2 : TROUVER LE VOL VERS LE PAYS LE MOINS CHER
' =============================================
Sub TrouverVolMoinsCher()

    ' On récupère la feuille des résultats
    Dim wsR As Worksheet
    Set wsR = ThisWorkbook.Sheets("Resultats")

    ' On cherche la dernière ligne remplie
    Dim derniereLigne As Long
    derniereLigne = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).row

    ' Si aucun produit n'a été trouvé, on avertit l'utilisateur
    If derniereLigne < 3 Then
        MsgBox "Effectuez d'abord une recherche.", vbExclamation
        Exit Sub
    End If

    ' On initialise les variables pour mémoriser le moins cher trouvé
    Dim prixMin    As Double: prixMin = 999999
    Dim paysMin    As String: paysMin = ""
    Dim produitMin As String: produitMin = ""
    Dim i          As Long
    Dim cellePrix  As String
    Dim valNum     As Double

    ' On parcourt toutes les lignes de résultats
    For i = 3 To derniereLigne

        ' On lit la valeur de la colonne Prix (colonne I)
        cellePrix = Trim(wsR.Cells(i, 9).Value)

        ' On ignore les lignes sans prix (affichées avec un tiret)
        If cellePrix <> "-" And cellePrix <> "" Then

            ' On convertit "1.99 EUR" en nombre 1.99
            valNum = Val(Replace(Replace(cellePrix, ",", "."), " EUR", ""))

            ' Si c'est le prix le plus bas trouvé jusqu'ici, on le mémorise
            If valNum > 0 And valNum < prixMin Then
                prixMin = valNum
                paysMin = wsR.Cells(i, 6).Value    ' Colonne F = Pays
                produitMin = wsR.Cells(i, 3).Value ' Colonne C = Nom du produit
            End If

        End If

    Next i

    ' Si aucun prix n'a été trouvé dans les résultats
    If paysMin = "" Then
        MsgBox "Aucun prix trouvé pour déterminer une destination.", vbExclamation
        Exit Sub
    End If

    ' Si le champ Pays contient plusieurs pays séparés par une virgule,
    ' on prend uniquement le premier
    If InStr(paysMin, ",") > 0 Then
        paysMin = Trim(Split(paysMin, ",")(0))
    End If

    ' On affiche un résumé et on demande confirmation à l'utilisateur
    Dim rep As Integer
    rep = MsgBox("Prix minimum trouvé : " & Format(prixMin, "0.00") & " EUR" & vbCrLf & _
                 "Produit : " & produitMin & vbCrLf & _
                 "Pays : " & paysMin & vbCrLf & vbCrLf & _
                 "Ouvrir Google Flights vers " & paysMin & " ?", _
                 vbYesNo + vbQuestion, "Trouver un vol")

    ' Si l'utilisateur clique Oui, on ouvre Google Flights dans le navigateur
    If rep = vbYes Then
        Shell "cmd /c start https://www.google.com/travel/flights?q=Flights+to+" & EncodeURL(paysMin), vbHide
    End If

End Sub


' =============================================
' TRADUIRE UN MOT EN PLUSIEURS LANGUES
' On utilise Google Translate pour ça
' =============================================
Function TraduireMultiLangues(mot As String) As String()

    ' Les 6 langues dans lesquelles on va traduire
    Dim langues As Variant
    langues = Array("en", "es", "de", "it", "pt", "nl")

    ' Le tableau qui contiendra toutes les traductions uniques
    Dim tousTermes(6) As String
    tousTermes(0) = LCase(mot) ' Le mot original en français est toujours le premier

    ' On crée un objet HTTP pour appeler Google Translate
    Dim http    As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim i       As Integer
    Dim reponse As String
    Dim debut   As Long
    Dim fin     As Long
    Dim trad    As String

    ' Pour chaque langue, on appelle Google Translate et on récupère la traduction
    For i = 0 To UBound(langues)

        On Error Resume Next
        http.Open "GET", "https://translate.googleapis.com/translate_a/single?client=gtx&sl=fr&tl=" & _
                  langues(i) & "&dt=t&q=" & EncodeURL(mot), False
        http.setRequestHeader "User-Agent", "ExcelOFFApp/1.0 (contact@monapp.com)"
        http.Send

        trad = ""
        If http.Status = 200 Then
            ' La traduction se trouve entre les deux premiers guillemets de la réponse
            reponse = http.responseText
            debut = InStr(reponse, """") + 1
            fin = InStr(debut, reponse, """")
            trad = LCase(Mid(reponse, debut, fin - debut))
        End If
        On Error GoTo 0

        ' On n'ajoute la traduction que si elle est différente de celles déjà trouvées
        ' (évite les doublons si deux langues donnent la même traduction)
        If trad = "" Or trad = LCase(mot) Then
            tousTermes(i + 1) = ""
        ElseIf InStr("|" & Join(tousTermes, "|") & "|", "|" & trad & "|") > 0 Then
            tousTermes(i + 1) = "" ' Doublon détecté, on l'ignore
        Else
            tousTermes(i + 1) = trad
        End If

        ' Petite pause entre chaque appel pour ne pas surcharger Google
        Wait 300

    Next i

    Set http = Nothing

    ' On reconstruit un tableau propre en retirant les cases vides
    Dim res() As String
    Dim n     As Integer: n = 0

    For i = 0 To 6
        If tousTermes(i) <> "" Then
            ReDim Preserve res(n)
            res(n) = tousTermes(i)
            n = n + 1
        End If
    Next i

    ' Sécurité : si aucune traduction n'a fonctionné, on garde au moins le mot français
    If n = 0 Then
        ReDim res(0)
        res(0) = LCase(mot)
    End If

    TraduireMultiLangues = res

End Function


' =============================================
' LIRE LA VALEUR D'UN CHAMP DANS UN JSON
' Exemple : Valeur(json, "product_name") retourne "Nutella"
' =============================================
Function Valeur(json As String, cle As String) As String

    Valeur = ""

    ' On cherche la position du nom du champ dans le JSON
    Dim pos As Long
    pos = InStr(json, """" & cle & """")
    If pos = 0 Then Exit Function ' Le champ n'existe pas dans ce JSON

    ' On avance jusqu'au ":" qui sépare le nom du champ de sa valeur
    pos = InStr(pos + Len(cle) + 1, json, ":") + 1

    ' On saute les espaces et le guillemet ouvrant si c'est une valeur textuelle
    Do While pos < Len(json) And (Mid(json, pos, 1) = " " Or Mid(json, pos, 1) = """")
        pos = pos + 1
    Loop

    Dim fin As Long

    ' Si la valeur est du texte (entre guillemets), on lit jusqu'au guillemet fermant
    If Mid(json, pos - 1, 1) = """" Then
        fin = InStr(pos, json, """")
    Else
        ' Si c'est un nombre, on lit jusqu'au premier caractère non numérique
        fin = pos
        Do While fin <= Len(json) And InStr("0123456789.-", Mid(json, fin, 1)) > 0
            fin = fin + 1
        Loop
    End If

    If fin > pos Then Valeur = Mid(json, pos, fin - pos)

End Function


' =============================================
' LIRE UN CHAMP JSON EN MODE TEXTE PUR
' Utilisé pour les codes-barres car ce sont de très grands
' nombres que VBA tronquerait en les traitant comme des chiffres
' =============================================
Function ValeurTexte(json As String, cle As String) As String

    ValeurTexte = ""

    Dim pos As Long
    pos = InStr(json, """" & cle & """")
    If pos = 0 Then Exit Function

    pos = InStr(pos + Len(cle) + 1, json, ":") + 1

    ' On saute les espaces
    Do While pos < Len(json) And Mid(json, pos, 1) = " "
        pos = pos + 1
    Loop

    Dim fin As Long
    Dim c   As String
    c = Mid(json, pos, 1)

    ' Si la valeur est entre guillemets, on lit jusqu'au guillemet fermant
    If c = """" Then
        pos = pos + 1
        fin = InStr(pos, json, """")
        If fin > pos Then ValeurTexte = Mid(json, pos, fin - pos)
    Else
        ' Sinon on lit jusqu'au prochain séparateur JSON
        fin = pos
        Do While fin <= Len(json) And InStr(",} " & Chr(13) & Chr(10), Mid(json, fin, 1)) = 0
            fin = fin + 1
        Loop
        If fin > pos Then ValeurTexte = Mid(json, pos, fin - pos)
    End If

End Function


' =============================================
' DÉCOUPER UN JSON EN LISTE DE PRODUITS INDIVIDUELS
' Le JSON contient un grand tableau [{produit1},{produit2}...]
' Cette fonction le découpe en morceaux, un par produit
' =============================================
Function SplitObjets(json As String) As String()

    Dim res()      As String: ReDim res(0)
    Dim i          As Long
    Dim depthAcc   As Integer  ' Compte les accolades ouvertes { }
    Dim depthCro   As Integer  ' Compte les crochets ouverts [ ]
    Dim dansChaine As Boolean  ' Sommes-nous à l'intérieur d'un texte entre guillemets ?
    Dim debut      As Long     ' Position du début du produit en cours
    Dim n          As Integer: n = 0
    Dim c          As String

    ' On parcourt le JSON caractère par caractère
    For i = 1 To Len(json)
        c = Mid(json, i, 1)

        ' Si on rencontre un guillemet non échappé, on entre ou sort d'une chaîne de texte
        ' Cela nous permet d'ignorer les { } et [ ] qui sont dans des valeurs textuelles
        If c = """" Then
            If i = 1 Or Mid(json, i - 1, 1) <> "\" Then
                dansChaine = Not dansChaine
            End If
        End If

        ' On ne traite les symboles structurels que si on n'est pas dans un texte
        If Not dansChaine Then
            Select Case c
                Case "{"
                    ' Si on ouvre une accolade au niveau racine, c'est le début d'un produit
                    If depthAcc = 0 And depthCro = 0 Then debut = i
                    depthAcc = depthAcc + 1
                Case "}"
                    depthAcc = depthAcc - 1
                    ' Si on revient à zéro, le produit est complet : on l'ajoute à la liste
                    If depthAcc = 0 And depthCro = 0 And debut > 0 Then
                        ReDim Preserve res(n)
                        res(n) = Mid(json, debut, i - debut + 1)
                        n = n + 1
                    End If
                Case "["
                    depthCro = depthCro + 1
                Case "]"
                    depthCro = depthCro - 1
            End Select
        End If

    Next i

    If n = 0 Then ReDim res(0)
    SplitObjets = res

End Function


' =============================================
' EXTRAIRE UN SOUS-TABLEAU D'UN JSON
' Exemple : ExtraireTableau(json, "products") retourne
' le contenu du tableau products : [{...},{...},...]
' =============================================
Function ExtraireTableau(json As String, cle As String) As String

    ' On cherche le nom du tableau dans le JSON
    Dim debut As Long
    debut = InStr(json, """" & cle & """")
    If debut = 0 Then Exit Function

    ' On avance jusqu'au crochet ouvrant [
    debut = InStr(debut, json, "[") + 1

    ' On parcourt le JSON en comptant les crochets pour trouver la fin du tableau
    Dim i     As Long
    Dim depth As Integer: depth = 1

    For i = debut To Len(json)
        If Mid(json, i, 1) = "[" Then depth = depth + 1
        If Mid(json, i, 1) = "]" Then depth = depth - 1
        ' Quand on retombe à zéro, on a trouvé la fin du tableau
        If depth = 0 Then
            ExtraireTableau = Mid(json, debut, i - debut)
            Exit Function
        End If
    Next i

End Function


' =============================================
' ÉCRIRE LES INFORMATIONS D'UN PRODUIT SUR UNE LIGNE
' =============================================
Sub AjouterLigne(wsR As Worksheet, ligne As Long, p As String, code As String)

    ' Colonne A : numéro d'ordre du produit dans la liste
    wsR.Cells(ligne, 1).Value = ligne - 2

    ' Colonne B : lien cliquable vers la photo du produit
    Dim urlPhoto As String
    urlPhoto = Valeur(p, "image_url")
    If urlPhoto <> "" Then
        wsR.Hyperlinks.Add wsR.Cells(ligne, 2), urlPhoto, , "Voir la photo", "Photo"
    End If

    ' Colonnes C à F : informations principales du produit
    wsR.Cells(ligne, 3).Value = CleanTxt(Valeur(p, "product_name")) ' Nom
    wsR.Cells(ligne, 4).Value = CleanTxt(Valeur(p, "brands"))       ' Marque
    wsR.Cells(ligne, 5).Value = CleanTxt(Valeur(p, "quantity"))     ' Quantité
    wsR.Cells(ligne, 6).Value = CleanTxt(Valeur(p, "countries"))    ' Pays de vente

    ' Colonnes G et H : scores nutritionnel (Nutri-Score) et écologique (Eco-Score)
    ' Ces scores vont de A (meilleur) à E (moins bon)
    Dim nutri As String: nutri = LCase(Valeur(p, "nutriscore_grade"))
    Dim eco   As String: eco = LCase(Valeur(p, "ecoscore_grade"))

    ' On n'affiche que les lettres valides (a à e), sinon on met "?" si inconnu
    wsR.Cells(ligne, 7).Value = IIf(nutri Like "[a-e]", UCase(nutri), "?")
    wsR.Cells(ligne, 8).Value = IIf(eco Like "[a-e]", UCase(eco), "?")

    ' Colonne J : code-barres, utilisé ensuite pour récupérer le prix
    wsR.Cells(ligne, 10).Value = code

    ' On colorie les cellules de score : vert pour A, rouge pour E
    ColorScore wsR.Cells(ligne, 7)
    ColorScore wsR.Cells(ligne, 8)

End Sub


' =============================================
' CHARGER LES PRIX DEPUIS OPENPRICES
' Pour chaque produit trouvé, on interroge l'API OpenPrices
' avec son code-barres pour récupérer son prix en magasin
' =============================================
Sub ChargePrix(wsR As Worksheet, derniereLigne As Long, ws As Worksheet)

    Dim r        As Long
    Dim code     As String
    Dim reponse  As String
    Dim tabItems As String
    Dim objets() As String

    For r = 3 To derniereLigne

        ' On lit le code-barres du produit en colonne J
        code = Trim(CStr(wsR.Cells(r, 10).Value))

        ' On appelle OpenPrices avec ce code-barres
        reponse = GetHTTP(URL_PRIX & "?product_code=" & code & "&page_size=1")

        ' On attend 300ms entre chaque appel pour ne pas surcharger le serveur
        Wait 300

        ' Si la réponse contient un prix, on l'extrait et on l'affiche
        If InStr(reponse, """price""") > 0 Then
            tabItems = ExtraireTableau(reponse, "items")
            If tabItems <> "" Then
                objets = SplitObjets(tabItems)
                If UBound(objets) >= 0 Then
                    ' On affiche le prix suivi de la devise, ex : "2.49 EUR"
                    wsR.Cells(r, 9).Value = Valeur(objets(0), "price") & " " & Valeur(objets(0), "currency")
                End If
            End If
        Else
            ' Pas de prix trouvé pour ce produit
            wsR.Cells(r, 9).Value = "-"
        End If

        ' Toutes les 5 lignes, on met à jour le message de progression
        If r Mod 5 = 0 Then
            Statut ws, "Chargement des prix : " & (r - 2) & "/" & (derniereLigne - 2)
        End If

    Next r

    ' Message final avec le nombre total de produits affichés
    Statut ws, "Terminé ! " & (derniereLigne - 2) & " produits trouvés."

End Sub


' =============================================
' ENVOYER UNE REQUÊTE HTTP ET RÉCUPÉRER LA RÉPONSE
' C'est la fonction qui "parle" au serveur OpenFoodFacts
' =============================================
Function GetHTTP(url As String) As String

    On Error Resume Next

    ' On crée un objet qui sait envoyer des requêtes sur internet
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' On prépare la requête (GET = on demande des données, on n'en envoie pas)
    http.Open "GET", url, False

    ' On s'identifie auprès du serveur (obligatoire pour ne pas être bloqué)
    http.setRequestHeader "User-Agent", "ExcelOFFApp/1.0 (contact@monapp.com)"

    ' On définit des délais maximum pour ne pas rester bloqué si le serveur ne répond pas
    http.setTimeouts 5000, 15000, 15000, 15000

    ' On envoie la requête
    http.Send

    ' Si le serveur répond avec le code 200, tout s'est bien passé : on retourne la réponse
    If http.Status = 200 Then
        GetHTTP = http.responseText
    Else
        ' Sinon on affiche le code d'erreur dans la fenêtre de débogage VBA
        Debug.Print "Erreur HTTP " & http.Status & " sur : " & url
    End If

    ' On libère proprement la mémoire
    Set http = Nothing
    On Error GoTo 0

End Function


' =============================================
' FAIRE UNE PAUSE (en millisecondes)
' Utilisé pour ne pas surcharger les serveurs
' =============================================
Sub Wait(ms As Long)
    Dim t As Single
    t = Timer
    ' On attend que le temps demandé soit écoulé
    ' DoEvents permet à Excel de rester réactif pendant l'attente
    Do While (Timer - t) * 1000 < ms
        DoEvents
    Loop
End Sub


' =============================================
' ÉCRIRE LES TITRES DE COLONNES EN LIGNE 2
' =============================================
Sub EcrireEntetes(wsR As Worksheet)

    Dim titres As Variant
    titres = Array("N°", "Photo", "Nom", "Marque", "Quantité", "Pays", "Nutri", "Eco", "Prix", "Code-barres")

    Dim i As Integer
    For i = 0 To 9
        With wsR.Cells(2, i + 1)
            .Value = titres(i)
            .Font.Bold = True
            .Interior.Color = RGB(46, 204, 113) ' Fond vert
            .Font.Color = vbWhite               ' Texte blanc
            .HorizontalAlignment = xlCenter
        End With
    Next i

End Sub


' =============================================
' AJUSTER LES LARGEURS DE COLONNES
' =============================================
Sub MiseEnForme(wsR As Worksheet)

    ' Largeurs définies pour chaque colonne (en nombre de caractères)
    Dim largeurs As Variant
    largeurs = Array(4, 8, 35, 20, 12, 15, 8, 8, 14, 15)

    Dim i As Integer
    For i = 0 To 9
        wsR.Columns(i + 1).ColumnWidth = largeurs(i)
    Next i

End Sub


' =============================================
' ENCODER UN TEXTE POUR L'UTILISER DANS UNE URL
' Les caractères spéciaux comme les espaces, accents...
' doivent être remplacés par leur code %XX dans une URL
' Exemple : "café" devient "caf%C3%A9"
' =============================================
Function EncodeURL(s As String) As String

    Dim i As Integer
    Dim c As String
    Dim r As String

    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        ' Les lettres et chiffres passent tels quels dans une URL
        If c Like "[A-Za-z0-9]" Then
            r = r & c
        Else
            ' Les autres caractères sont remplacés par %XX
            r = r & "%" & Right("0" & Hex(Asc(c)), 2)
        End If
    Next i

    EncodeURL = r

End Function


' =============================================
' REMPLACER LES CODES UNICODE PAR DE VRAIS CARACTÈRES
' Le JSON retourné par l'API contient parfois des codes
' comme \u00e9 à la place de é. Cette fonction les remplace.
' =============================================
Function CleanTxt(s As String) As String

    Dim r As String: r = s

    r = Replace(r, "\u00e9", "é")
    r = Replace(r, "\u00e8", "è")
    r = Replace(r, "\u00ea", "ê")
    r = Replace(r, "\u00eb", "ë")
    r = Replace(r, "\u00e0", "à")
    r = Replace(r, "\u00e2", "â")
    r = Replace(r, "\u00f4", "ô")
    r = Replace(r, "\u00fb", "û")
    r = Replace(r, "\u00f9", "ù")
    r = Replace(r, "\u00ee", "î")
    r = Replace(r, "\u00ef", "ï")
    r = Replace(r, "\u00e7", "ç")
    r = Replace(r, "\u2019", "'")
    r = Replace(r, "\u00ab", "«")
    r = Replace(r, "\u00bb", "»")
    r = Replace(r, "\u00e6", "æ")
    r = Replace(r, "\u0153", "œ")

    CleanTxt = r

End Function


' =============================================
' COLORISER UNE CELLULE SELON LA LETTRE DU SCORE
' A = vert foncé (très bien)
' B = vert clair (bien)
' C = jaune (moyen)
' D = orange (mauvais)
' E = rouge (très mauvais)
' ? = gris (inconnu)
' =============================================
Sub ColorScore(c As Range)

    Select Case LCase(c.Value)
        Case "a": c.Interior.Color = RGB(0, 128, 0):   c.Font.Color = vbWhite
        Case "b": c.Interior.Color = RGB(144, 238, 144)
        Case "c": c.Interior.Color = RGB(255, 215, 0)
        Case "d": c.Interior.Color = RGB(243, 133, 33): c.Font.Color = vbWhite
        Case "e": c.Interior.Color = RGB(255, 0, 0):   c.Font.Color = vbWhite
        Case Else: c.Interior.Color = RGB(220, 220, 220)
    End Select

    c.HorizontalAlignment = xlCenter

End Sub


' =============================================
' AFFICHER UN MESSAGE DE PROGRESSION EN CELLULE D13
' Appelé régulièrement pendant l'exécution pour informer
' l'utilisateur de ce que fait le programme
' =============================================
Sub Statut(ws As Worksheet, msg As String)
    ws.Range("D13").Value = msg
    ' DoEvents laisse Excel respirer et rafraîchir l'écran pendant l'exécution
    DoEvents
End Sub
