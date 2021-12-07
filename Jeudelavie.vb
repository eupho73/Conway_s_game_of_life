'  ~~~~~~~~~~~~~~~ Code de la Feuil1  ~~~~~~~~~~~~~~~~~~ 


Option Explicit


Public Sub nextG(ligne, colonne) 'Procèdure qui renvoie l'état d'une cellule à la prochaine génération, en fonctions des règles choisies
    Dim compteur As Integer
    Dim i As Integer
    Dim j As Integer
    compteur = 0
    For i = ligne - 1 To ligne + 1 'boucle qui relève les valeurs des cases adjacentes en comptant elle-même
        For j = colonne - 1 To colonne + 1
            If i = 0 Then i = 50  'Permet de gérer les conditions de bord : tout ce qui se trouve en dehors de grille ne contient pas de cellule
            If i = 51 Then i = 1
            If j = 0 Then j = 70
            If j = 71 Then j = 1
            compteur = compteur + Worksheets("Plateau").Cells(i, j)
        Next
    Next
    compteur = compteur - Worksheets("Plateau").Cells(ligne, colonne).Value  'si la case est vivante on soustrait 1 au compteur

    If Worksheets("Plateau").Cells(ligne, colonne).Value = 0 Then  'conditions de naissance selon les règles choisies dans Options
            If Options.CheckBox1.Value = True And compteur = 1 _
                Or Options.CheckBox2.Value = True And compteur = 2 _
                Or Options.CheckBox3.Value = True And compteur = 3 _
                Or Options.CheckBox4.Value = True And compteur = 4 _
                Or Options.CheckBox5.Value = True And compteur = 5 _
                Or Options.CheckBox6.Value = True And compteur = 6 _
                Or Options.CheckBox7.Value = True And compteur = 7 _
                Or Options.CheckBox8.Value = True And compteur = 8 Then
                Worksheets("nextGen").Cells(ligne, colonne).Value = 1
                
            Else
                Worksheets("nextGen").Cells(ligne, colonne).Value = 0
            End If
        Else   'conditions de survie i.e. la case est vivante et vaut 1
            If Options.CheckBox10.Value = True And compteur = 1 _
                Or Options.CheckBox11.Value = True And compteur = 2 _
                Or Options.CheckBox12.Value = True And compteur = 3 _
                Or Options.CheckBox13.Value = True And compteur = 4 _
                Or Options.CheckBox14.Value = True And compteur = 5 _
                Or Options.CheckBox15.Value = True And compteur = 6 _
                Or Options.CheckBox16.Value = True And compteur = 7 _
                Or Options.CheckBox17.Value = True And compteur = 8 Then
                Worksheets("nextGen").Cells(ligne, colonne).Value = 1
            Else
                Worksheets("nextGen").Cells(ligne, colonne).Value = 0
            End If
        End If
    
End Sub

Private Sub CommandButton1_Click()  'C'est le bouton pour lancer/arrêter le jeu. C'est un activeX, c'est pour ça qu'il se trouve sur Feuil1, il ne peut pas être dans module
    If CommandButton1.Caption = "Arrêter" Then  'le jeu est en cours, on veut l'arrêter
    CommandButton1.Caption = "Jouer"
    CommandButton1.BackColor = &HFF00FF
    arret = True 'permet d'arrêter le jeu quand on clique sur le commandButton "jouer" de caption "Arrêter"
    
    ElseIf CommandButton1.Caption = "Jouer" Then  'on veut démarrer le jeu
    CommandButton1.Caption = "Arrêter"
    CommandButton1.BackColor = &HFF797C
    Call jouer  'On lance le jeu en appelant jouer
    End If
End Sub



Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)  'changer l'état d'une cellule avec clique droit
    
    Cancel = True
    If Not Intersect(Target, Range("grille")) Is Nothing Then  'On a nommé la plage (A1:BR50) grille dans le gestionnaire de nom
        Select Case Target.Value
            Case 0
                Target.Value = 1   'Si la cellule est morte, on la rend vivante et inversement
            Case 1
                Target.Value = 0
        End Select
    End If
    
End Sub




Private Sub Auto_Open()
    
    Options.CheckBox3.Value = True  'par défault, à l'ouverture du jeu, on met les règles S23/B3 et pas de limite d'itération
    Options.CheckBox11.Value = True
    Options.CheckBox12.Value = True
    Options.OptionButton2.Value = True
    UserForm1.modele = "glider"  'car il est True par défaut dans l'UserForm1
    
End Sub

Public Sub chargerTable(modele)  'modele est récupérée dans l'userform Librairie. chargerTable permet de charger le modele renseigné déjà existant.

        Worksheets(modele).Range("A1", "BR50").Copy
        Range("grille").PasteSpecial Paste:=xlPasteValues  'deux fois car ne fonctionne pas toujours en simple
        Range("grille").PasteSpecial Paste:=xlPasteValues
        Worksheets("Plateau").Range("A1").Select 'pour ne pas avoir tout le tableau selectionné juste après la copie

End Sub








'~~~~~~~~~~~~~~~~ Les Userforms  ~~~~~~~~~~~~~~~~~


'--------------------  L'UserForm Aleatoire

Option Explicit


Private Sub curseurRemplissageAleatoire_Change()  'action quand l'utilisateur modifie la valeur via le curseur

    probaTextBox = curseurRemplissageAleatoire.Value  'met le pourcentage choisi sur le curseur dans le textBox
    'probaTextBox est le nom du textbox contenant les valeurs affichées par le curseur
    
End Sub


Private Sub ValiderRemplissageAleatoire_Click()

    Call plateauAleatoire  'On appelle la procèdure, c'est elle qui remplit aléatoirement la grille
    aleatoire.Hide
    'On ne peut pas unload car on utilise curseurRemplissageAleatoire.Value
    Worksheets("Plateau").Cells(46, 72) = "Modèle : Aléatoire (" & probaTextBox & " %" & " de cellules vivantes)"  'On modifie le nom du modèle dans la partie statut
    
End Sub

Private Sub Aleatoire_Initialize()     'Propriétés du module au démarrage

    probaTextBox.Locked = True 'L'utilisateur ne peut pas modifier le nombre dans le txtBox
    With curseurRemplissageAleatoire
        .Min = 0
        .Max = 100
        .LargeChange = 5
        
End Sub

Public Function coinFlip() As Integer  'simule un schéma de Bernouilli de paramètre la valeur rentrée avec le curseur
    
    If Rnd() < (aleatoire.curseurRemplissageAleatoire.Value / 100) Then  'équivalent à x% d'être vivant
        coinFlip = 1
    Else
        coinFlip = 0
    End If
    
End Function


Sub plateauAleatoire()  
'remplit la grille aléatoirement selon la valeur rensignée
    
    Dim i, j As Integer
    For i = 1 To 50
        For j = 1 To 70
            Worksheets("Plateau").Cells(i, j).Value = coinFlip
        Next
    Next
    
End Sub





'----------------------- L'UserForm Options


Option Explicit
Public regleSurvie, regleNaissance As String



Private Sub CommandButton1_Click()  'informations sur le délai. Bouton "?"
    MsgBox "Le délai permet d'augmenter la durée entre chaque générations. La vitesse d'exécution du programme dépend des performances de votre ordinateur et de l'efficacité de l'algorithme. Le délai peut aller jusqu'à 10 secondes."
End Sub

Private Sub SpinButton1_Change()
    delaiTextBox = SpinButton1.Value   'on lie la toupie à la textbox
End Sub

Private Sub delaiTextBox_Change()
    SpinButton1.Value = Val(delaiTextBox)   'on lit les valeurs rentrées manuellement dans la textbox à la toupie
End Sub

Private Sub validerOptions_Click()
    regleNaissance = ""  'On veut savoir quelles règles sont actives sur le plateau sans avoir à cliquer sur options
    If CheckBox1.Value = True Then regleNaissance = regleNaissance & "1"
    If CheckBox2.Value = True Then regleNaissance = regleNaissance & "2"
    If CheckBox3.Value = True Then regleNaissance = regleNaissance & "3"
    If CheckBox4.Value = True Then regleNaissance = regleNaissance & "4"
    If CheckBox5.Value = True Then regleNaissance = regleNaissance & "5"
    If CheckBox6.Value = True Then regleNaissance = regleNaissance & "6"
    If CheckBox7.Value = True Then regleNaissance = regleNaissance & "7"
    If CheckBox8.Value = True Then regleNaissance = regleNaissance & "8"
    
    regleSurvie = ""
    If CheckBox10.Value = True Then regleSurvie = regleSurvie & "1"
    If CheckBox11.Value = True Then regleSurvie = regleSurvie & "2"
    If CheckBox12.Value = True Then regleSurvie = regleSurvie & "3"
    If CheckBox13.Value = True Then regleSurvie = regleSurvie & "4"
    If CheckBox14.Value = True Then regleSurvie = regleSurvie & "5"
    If CheckBox15.Value = True Then regleSurvie = regleSurvie & "6"
    If CheckBox16.Value = True Then regleSurvie = regleSurvie & "7"
    If CheckBox17.Value = True Then regleSurvie = regleSurvie & "8"
    'regleNaissance et regleSurvie sont donc des chaines de caractères formées d'une suite de chiffres
    
    Worksheets("Plateau").Cells(45, 72).Value = "Règles : " & "S" & regleSurvie & "/B" & regleNaissance  'Je remplit le statut avec les nouvelle règles
    Worksheets("Plateau").Cells(43, 72).Value = "Délai : " & SpinButton1.Value & " ms"
    Options.Hide 'On ne peut pas unload car on utilise directement les valeurs des CheckBox dans nextG
    'On préfère utiliser directement les valeurs ainsi car cela permet d'économiser en variable.
    'On aurait pu stocké les choix de règles et d'options dans des variables créés à cet effet mais cela aurait demandé au moins
    '16 variables pour les règles et une modificatin de la fonction nextG de la Feuil1
    
End Sub


Private Sub annuler_Click()  'restaure les règles par défault (S23/B3)
  Dim oCtrl As Control
    For Each oCtrl In Me.Controls    'Au lieu de décocher chaque CheckBox en invoquant son nom, je cherche tous les contrôles de type CheckBox
      If TypeOf oCtrl Is MSForms.CheckBox Then
        oCtrl.Value = False
      End If
    Next
    CheckBox11.Value = True
    CheckBox12.Value = True
    CheckBox3.Value = True
    iterationsOpt.Value = False
    OptionButton2.Value = True
    
End Sub


Private Sub Options_Initialize()  'On initialise les règles S23/B3 par défaut

    CheckBox11.Value = True
    CheckBox12.Value = True
    CheckBox3.Value = True
    
End Sub




'------------------- L'UserForm Sauvegarde



Private Sub validerSauvegarde_Click()
    
    Dim iRow As Integer   'le prochain indice de ligne de notre liste de noms sauvegardés qui n'est pas remplie

    iRow = Worksheets("Plateau").Range("BW" & Rows.Count).End(xlUp).Row + 1 'Je compte le nombre de ligne déjà présent dans la plage nommée Sauvegardes
    On Error GoTo ErrMsg  'si jamais le nom n'est pas valide (nom vide, caractère non permit)
    
    If nomSauvegarde.Text = "Plateau" Or nomSauvegarde.Text = "nextGen" _
        Or nomSauvegarde.Text = "heavyweightemulator" Or nomSauvegarde.Text = "gourmet" _
        Or nomSauvegarde.Text = "glider" Or nomSauvegarde.Text = "gliders12" _
        Or nomSauvegarde.Text = "blinkerfuse" Or nomSauvegarde.Text = "glidergun" _
        Or nomSauvegarde.Text = "sauvegarde" Or nomSauvegarde.Text = "Sauvegardes" Then GoTo ErrMsg2
    'éviter un problème dans l'appel d'un modèle sont le nom existerait en double dans la librairie
    
    
    
    Worksheets("Plateau").Cells(iRow, 75).Value = nomSauvegarde.Text    'je stocke le nom du nouveau modèle dans le tableau Sauvegardes. Il se place directement à la suite (utilisation de iRow)
    Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = nomSauvegarde.Text   'on crée une nouvelle feuille avec le nom inscrit dans la textBox
    Worksheets("Plateau").Range("A1", "BR50").Copy
    Worksheets(nomSauvegarde.Text).Range("A1", "BR50").PasteSpecial Paste:=xlPasteValues  'et on copie dessus les valeurs qu'on veut sauvegarder
    Worksheets("Plateau").Activate 'Je reviens sur Plateau
    Worksheets("Plateau").Cells(1, 1).Select
    Worksheets(nomSauvegarde.Text).Visible = False   'On cache la feuil à l'utilisateur. On aurait préféré la rendre veryhidden mais cela pose des problèmes lorsque 
    'l'on veut faire certaines action sur cette feuille, comme par exemple la supprimer à partir du module Charger
    
    MsgBox ("Sauvegarde réussie !")
    Sauvegarder.Hide
    
    Exit Sub
    
ErrMsg:
    MsgBox ("Veuillez rentrer un nom valide.")
    Application.DisplayAlerts = False   'On ne veux pas montrer qu'on delete une feuille
    Worksheets(Worksheets.Count).Delete   'On delete car en cas d'erreur, excel crée une feuille de type FeuilN (cette feuille se truve bien à la fin)

ErrMsg2:
    MsgBox ("Ce modèle existe déjà dans la bibliothèque des modèles préenregistrés. Veuillez rentrer un autre nom.")

End Sub




'------------------------- L'UserForm Librairie (UserForm1)


Option Explicit

Public modele As String


Private Sub blinkerfuseOpt_Click()   'Les modèles préenregistrés 

    If blinkerfuseOpt.Value = True Then UserForm1.modele = "blinkerfuse" 'On assigne à la valeur globale modele le nom du modele qui cïncide avec le nom de la feuille
    
End Sub

Private Sub gliderOpt_Click()

    If gliderOpt.Value = True Then UserForm1.modele = "glider"
    
End Sub

Private Sub gliders12Opt_Click()

    If gliders12Opt.Value = True Then UserForm1.modele = "gliders12"

End Sub

Private Sub gourmetOpt_Click()

    If gourmetOpt.Value = True Then UserForm1.modele = "gourmet"

End Sub

Private Sub heavyOpt_Click()

    If heavyOpt.Value = True Then UserForm1.modele = "heavy"

End Sub

Private Sub glidergunOpt_Click()

    If glidergunOpt.Value = True Then UserForm1.modele = "glidergun"
    
End Sub


Private Sub Charger_Click()    'Quand on clique sur le bouton valider

    Feuil1.chargerTable (UserForm1.modele)   'Je charge le modèle choisi avec la procédure chargerTable
    Worksheets("Plateau").Cells(47, 72).Value = "Modèle : " & modele  'Je modifie le nom du modèle choisi dans le statut du jeu
    UserForm1.Hide
    Unload UserForm1

End Sub


Private Sub chargerSauvegarde_Click()   'Quand on clique sur le bouton charger de la page des modèles sauvegardés
    
    modele = ListBox1.Text   'je récupère ce qui est séléctionné par l'utilisateur dans la listBox contenant les différents modèles sauvegardéss
    Feuil1.chargerTable (UserForm1.modele)   'Je charge ce modèle
    Worksheets("Plateau").Cells(47, 72).Value = "Modèle : " & modele   'Je modifie le nom du modèle choisi dans le statut du jeu
    UserForm1.Hide
    
End Sub

Private Sub enableSupprimer() 'on ne peut pas supprimer une sauvegarde si c'est la seule restante (ça provoquerait une erreur sur la formule de la plage Sauvegardes)
    
    If Worksheets("Plateau").Range("Sauvegardes").Rows.Count = 1 Then   'Si je n'ai qu'une ligne restante dans ma plage Sauvegardes
        UserForm1.supprimerSauvegarde.Enabled = False   'J'empêche la séléction du bouton "Supprimer"
    Else
        UserForm1.supprimerSauvegarde.Enabled = True   'Sinon, tout va bien
    End If
    
End Sub


Private Sub ListBox1_Click()    'Action quand je clique sur un élément de la listBox
    'Utile lorsque l'on veut supprimer tout les éléments un par un dans le module. (Il existe la même action lors du click sur le bouton "Charger", voir plus loin)
    
    Call enableSupprimer    'si >1 élément on peut en supprimer sinon on ne peut pas

End Sub

Private Sub supprimerSauvegarde_Click()   'permet de supprimer une sauvegarde séléctionnée dans la listbox
    
    Dim i As Long, finalRow As Long
    Dim nom As String
    Dim k As Integer
    
    finalRow = Worksheets("Plateau").Cells(Rows.Count, 75).End(xlUp).Row   'je compte le nombre de ligne dans Sauvegardes, i.e. le nombre de sauvegardes. 75 correspond à l'indice de colonne de la plage Sauvegardes
    nom = ListBox1.Text
    With Worksheets("Plateau")
    For i = finalRow To 71 Step -1  'Je boucle de la dernière ligne de la plage Sauvegardes (cette plage est dynamique) à sa première ligne (71)
        If Range("BW" & i).Value = nom Then
            If i <> 71 Then
                Range("BW" & i).EntireRow.Delete   'En supprimant toute la ligne, les lignes suivantes remontent automatiquement et donc leur indice de ligne diminue de 1
            Else     'Si je supprime le premier éléement du groupe Sauvegardes (ligne 71), j'ai des problèmes de références sur le groupe
                For k = 71 To finalRow - 1 Step 1
                    Worksheets("Plateau").Cells(k, 75).Value = Worksheets("Plateau").Cells(k + 1, 75).Value 'J'avance tout les noms d'une case, de sorte à garder la référence tout en se débarrassant du nom
                Next k
                Worksheets("Plateau").Range("BW" & finalRow).EntireRow.Delete  'Je peux donc supprimer le doublon en dernière ligne
            End If
            
        End If
    Next i
    End With
    Application.DisplayAlerts = False 'Pur ne pas voir le message d'avertissement quand on supprime une feuille
    Worksheets(nom).Delete
    
    Exit Sub
     
End Sub





'---------------------- Le code du Module 1


#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)  'Permet d'ajouter du délai
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Option Explicit

Public different As Boolean
Public arret As Boolean
Public generationN As Integer



Private Sub enableSupprimer() 'on ne peut pas supprimer une sauvegarde si c'est la seule restante

        If UserForm1.ListBox1.ListCount = 1 Then
                UserForm1.supprimerSauvegarde.Enabled = False
        Else
                UserForm1.supprimerSauvegarde.Enabled = True
        End If

End Sub

Private Sub Bouton1_Cliquer()  'Permet d'ouvrir la libraiie

    Call enableSupprimer  'On vérifie avant si il y a plus d'un élément sauvegardé (si oui, le bouton supprimé est enabled)
    UserForm1.Show

End Sub

Private Sub Effacer_Cliquer()    'permet de rendre toute les cases mortes

    Worksheets("Plateau").Range("A1", "BR50") = 0
    Worksheets("nextGen").Range("A1", "BR50") = 0
    Worksheets("Plateau").Cells(46, 72).Value = "Modèle : blank"
    Worksheets("Plateau").Cells(42, 72).Value = "Génération n° 0"

End Sub

Private Sub Remplir_Cliquer()   'Toutes les case vivantes

    Worksheets("Plateau").Range("grille") = 1
    Worksheets("nextGen").Range("A1", "BR50") = 1

End Sub

Private Sub Instructions_Cliquer()  'Affiche les instructions

    MsgBox ("Tu peux changer l'état d'une cellule en faisant clique droit dessus. Une cellule morte est représentée en blanc tandis qu'une cellule vivante est représentée en noir." & Chr(13) & Chr(10) _
    & "Une cellule morte à la génération n devient vivante à la génération suivante si exactement trois cellules voisines sont vivantes à la génération n, sinon elle reste morte." _
    & "Une cellule vivante à la génération n reste vivante à la génération n+1 si deux ou trois cellules voisines sont vivantes à la génération n, sinon elle meurt." _
    & " Il est possible de changer les règles d'évolution des cellules dans Options." & Chr(13) & Chr(10) _
    & "Le bouton charger permet d'accèder à des modèles préenregistrés et à tes modèles sauvegardés." & Chr(13) & Chr(10) & Chr(13) & Chr(10) _
    & "Amuses-toi bien !"), Title:="Instructions"

End Sub

Private Sub Aleatoire_Cliquer()  'Affiche le module Aleatoire

    aleatoire.Show

End Sub

Private Sub Options_Cliquer()  'Affiche le module Options

    Options.Show

End Sub

Private Sub Sauvegarder_Cliquer()   'Affiche le module Sauvegarde

    Sauvegarder.Show
    
End Sub


Private Sub test()   'Permet de calculer la prochaine génération et de l'afficher sur la grille de la feuille Plateau.
    
    Dim ligne As Integer
    Dim col As Integer
    For ligne = 1 To 49
        For col = 1 To 69
            Call Feuil1.nextG(ligne, col)  'Je remplis les cellules dans la feuille nextGen
            If Worksheets("nextGen").Cells(ligne, col) <> Worksheets("Plateau").Cells(ligne, col) Then different = True
            'Dès qu'une des cellule de nextGen est différente de celle de Plateau, different = True (la boucle ne s'arrête pas)
        Next
    Next
    If different = False Then arret = True 'On initialise different = False à chaque call de test() dans jouer()
    'Si je n'ai pas trouvé de changement entre la génération n et n+1, different va rester faux dans les boucles de lignes et de colonnes au dessus
    ' et donc va arrêter le jeu une fois arrivé à ce test.
    
    Worksheets("nextGen").Range("A1", "BR50").Copy   'L'utilisateur ne voit que Plateau, donc je récupère la grille de génération n+1 sur nextGen et je 
    'la met sur la grille de Plateau. On vient de passer de la génération n à la génération n+1 pour l'utilisateur
    Worksheets("Plateau").Range("A1", "BR50").PasteSpecial Paste:=xlPasteValues
    Worksheets("Plateau").Range("A1").Select
            
End Sub

Public Sub jouer()   'Permet de calculer les prochaines générations en bouclant en fonction des paramètres choisis par l'utilisateur

    arret = False   
    generationN = 0
    If Options.OptionButton2.Value = True Then 'cas sans limite d'itérations

        While Not arret
            Worksheets("Plateau").Cells(41, 72).Value = "Génération n° " & generationN   'génération n dans le statut de la simulation
            different = False   'On initialise pour la procédure test()
            Application.ScreenUpdating = False  'smooth transitions between states, permet notamment d'interagir pendant la simulation
            Application.EnableEvents = False
            Call test   'L'utilisateur voit la nouvelle génération
            generationN = generationN + 1   'On incrémente
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            DoEvents
            Sleep (Options.delaiTextBox)  'delai que l'utilisateur a rentré
            
        Wend

        If Not different Then MsgBox ("Forme stable")   'Il n'y a pas de différences avec la génération suivante, on peut arrêter le programme
        'test va remarquer qu'il n'y a pas de difference et va assigner à arret la valeur True
        Feuil1.CommandButton1.Caption = "Jouer"  'Le programme est arrêté -> on ne propose pas l'option arrêt mais l'option jouer
        Feuil1.CommandButton1.BackColor = &HFF00FF

    End If

    If Options.iterationsOpt.Value = True Then 'cas avec limite d'itération

        Dim i As Integer

        For i = 0 To Options.iterationTextBox  'limite imposée par l'utilisateur

            If Not arret Then  'On fait la même chose que dans le cas sans limite d'itération

                Worksheets("Plateau").Cells(41, 72).Value = "Génération n° " & i & " / " & Options.iterationTextBox
                Application.ScreenUpdating = False  'smooth transitions between states
                Application.EnableEvents = False
                Call test
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                DoEvents
                Sleep (Options.delaiTextBox)

            End If

            If Not different Then  'le programme s'arrête
                MsgBox ("Forme stable")
                Feuil1.CommandButton1.Caption = "Jouer"  'Le programme est arrêté -> on ne propose pas l'option arrêt mais l'option jouer
                Feuil1.CommandButton1.BackColor = &HFF00FF
                Exit For

            End If

        Next i

        Feuil1.CommandButton1.Caption = "Jouer"  'Le programme est arrêté -> on ne propose pas l'option arrêt mais l'option jouer
        Feuil1.CommandButton1.BackColor = &HFF00FF

    End If
            
End Sub
