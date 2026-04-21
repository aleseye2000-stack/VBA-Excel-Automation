#  VBA Excel – Outil Avancé de Comparaison de Données de Facturation

##  Projet réalisé chez Orange Business

##  Contexte

Dans le cadre de mes missions chez Orange Business, j’ai développé un outil VBA avancé permettant de comparer automatiquement des données de facturation entre différentes sources Excel.

Ce projet répond à un besoin métier critique : fiabiliser et accélérer la détection d’écarts dans des fichiers volumineux, initialement traités manuellement.

---

##  Objectifs

* Automatiser la comparaison de données entre deux ensembles Excel
* Identifier rapidement les écarts de montants
* Gérer des cas complexes (cellules vides, texte, formats différents)
* Fiabiliser les contrôles de facturation
* Réduire significativement le temps de traitement

---

##  Fonctionnalités développées

###  1. Comparaison intelligente des données

* Comparaison cellule par cellule entre deux colonnes/plages
* Gestion des décalages via `Offset()`
* Vérification de cohérence des tailles de plages

---

###  2. Mise en évidence automatique

* Coloration dynamique des écarts
* Distinction claire entre :

  * Valeurs identiques
  * Valeurs différentes
  * Valeurs non comparables

---

###  3. Gestion des types de données (robustesse)

* Détection des valeurs numériques avec `IsNumeric()`
* Prise en compte :

  * Cellules vides
  * Données texte
  * Formats hétérogènes

---

###  4. Sécurisation des entrées utilisateur

* Vérification du nombre de sélections (`Selection.Areas.Count`)
* Blocage des cas invalides
* Messages d’erreur explicites pour guider l’utilisateur

---

###  5. Fiabilisation des comparaisons

* Évitement des erreurs VBA (`Type mismatch`)
* Gestion des cas où les cellules ne correspondent pas correctement
* Contrôle des correspondances ligne par ligne

---

##  Code VBA (version représentative)

```vba id="zzs8f2"
Sub ComparerPlagesAvance()

    Dim cellule1 As Range
    Dim cellule2 As Range

    ' Vérification : une seule plage sélectionnée
    If Selection.Areas.Count > 1 Then
        MsgBox "Veuillez sélectionner une seule plage continue.", vbExclamation
        Exit Sub
    End If

    ' Parcours des cellules sélectionnées
    For Each cellule1 In Selection
        
        ' Correspondance avec la cellule adjacente (colonne de droite)
        Set cellule2 = cellule1.Offset(0, 1)
        
        ' Vérification que les deux cellules ne sont pas vides
        If cellule1.Value <> "" And cellule2.Value <> "" Then
            
            ' Vérification des types numériques
            If IsNumeric(cellule1.Value) And IsNumeric(cellule2.Value) Then
                
                ' Comparaison des valeurs
                If CDbl(cellule1.Value) <> CDbl(cellule2.Value) Then
                    cellule1.Interior.Color = vbRed
                    cellule2.Interior.Color = vbRed
                End If
                
            Else
                ' Cas non numérique → coloration différente
                cellule1.Interior.Color = vbYellow
                cellule2.Interior.Color = vbYellow
            End If
            
        End If
        
    Next cellule1

End Sub
```

---

##  Cas d’usage métier

* Comparaison de fichiers de facturation fournisseurs
* Vérification de cohérence entre systèmes (export vs reporting)
* Contrôle des montants avant validation financière
* Audit rapide de données Excel

---

##  Problèmes rencontrés et solutions

### ❌ 1. Erreur VBA "Type mismatch" (Run-time error 13)

**Cause :**

* Comparaison directe de cellules contenant texte, vide ou formats mixtes

✅ **Solution :**

* Utilisation de `IsNumeric()`
* Conversion sécurisée avec `CDbl()`

---

### ❌ 2. Bug avec sélection multiple

**Cause :**

* Le code ne gérait pas les sélections discontinues

✅ **Solution :**

* Vérification avec `Selection.Areas.Count`
* Restriction à une seule plage continue

---

### ❌ 3. Comparaisons incohérentes

**Cause :**

* Mauvaise correspondance entre cellules

✅ **Solution :**

* Utilisation de `Offset()` pour garantir l’alignement
* Comparaison ligne par ligne

---

### ❌ 4. Absence de gestion des cas non numériques

**Cause :**

* Certaines cellules contenaient du texte ou des valeurs inattendues

✅ **Solution :**

* Ajout d’un traitement spécifique avec coloration différente (jaune)

---

##  Améliorations possibles

* Interface utilisateur (UserForm) pour sélection dynamique
* Comparaison entre fichiers Excel différents
* Génération automatique d’un rapport des écarts
* Ajout de logs et historisation
* Optimisation des performances sur grands volumes

---

##  Impact métier

*  Réduction importante du temps de contrôle manuel
*  Diminution des erreurs humaines
*  Amélioration de la qualité des données financières
*  Processus de validation plus fiable et rapide

---

##  Fichiers inclus

* Fichier Excel macro-enabled (.xlsm)
* Module VBA exporté (.bas)
* Documentation du projet

---

## 👤 Auteur

Mame Alé SEYE
Étudiant ingénieur – Data & Intelligence Artificielle
