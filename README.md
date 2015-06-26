# VBA_Graph_In_Picture
Permet de passer en revue l'ensemble des graphiques d'un classeur pour en générer les images.

##Lien vers le site
http://customoffice.github.io/VBA_Graph_In_Picture/

## Instruction
- Soit créer un module dans votre projet vba et y copier/coller le code ci-dessous
- Soit télécharger le module (fichier *.bas) et l'inserer dans votre projet vba

##Code
```bash
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!TITRE : Génération pour le classeur, une image par graphique                                    !!!
'!!!DATE :  17.04.15                                                                                !!!
'!!!                                                                                                !!!
'!!!DESCRIPTION : permet de passer en revue l'ensemble des graphiques d'un classeur pour générer les!!!
'!!!images.                                                                                         !!!
'!!!                                                                                                !!!
'!!!REGLES :                                                                                        !!!
'!!!- fonctionne avec la macro creer_image                                                       !!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


Sub graph_in_picture()
    'Déclaration des variables
    Dim nom As String, feuille As String
    Dim i As Integer
    
    ScreenUpdating = False
    i = ActiveSheet.Index 'mémorise depuis quel onglet on lance la macro car creer_image sélectionne l'onglet
    
    'passe en revue tous les onglets d'un classeur
    For Each objet_feuille In ActiveWorkbook.Sheets
        feuille = objet_feuille.Name
        'passe en revue tous les graphiques d'un onglet
        For Each object_test In Worksheets(feuille).ChartObjects
            nom = object_test.Name
            Call creer_image(nom, feuille)
        Next
    Next
    
    Sheets(i).Select 'resélectionne l'onglet de départ
    
    ScreenUpdating = True
End Sub
```
