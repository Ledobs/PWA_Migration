Attribute VB_Name = "modDiagnosticOLE"
Option Explicit

' =============================================================================
' INVENTAIRE ET DIAGNOSTIC DES OBJETS OLE - MICROSOFT PROJECT DESKTOP
' =============================================================================
'
' Auteur             : François Breton
' Date de création   : 2026-08-06
'
' -----------------------------------------------------------------------------
' OBJECTIF
' -----------------------------------------------------------------------------
'
' Ce module permet aux pilotes système d'investiguer la présence d'objets OLE
' dans un plan Microsoft Project lorsqu'un avertissement de sécurité OLE est
' présenté à l'ouverture, à la consultation ou au traitement d'un projet.
'
' Le cas ayant mené à la création de cet outil concernait notamment des plans
' migrés d'un environnement Project Web App (PWA) vers un autre environnement.
' Certains éléments historiques contenus dans les plans pouvaient déclencher
' l'avertissement OLE sans qu'il soit évident de déterminer quelle tâche,
' ressource ou affectation en était responsable.
'
'
' -----------------------------------------------------------------------------
' ÉLÉMENTS ANALYSÉS
' -----------------------------------------------------------------------------
'
' Le diagnostic parcourt les principaux objets Project susceptibles de
' contenir, de référencer ou de transporter du contenu OLE :
'
'   - la tâche récapitulative du projet (Project Summary Task / tâche 0);
'   - les tâches;
'   - les ressources;
'   - les affectations;
'   - les objets OLE directs;
'   - les champs liés par OLE (LinkedFields);
'   - les Notes associées aux tâches, ressources et affectations.
'
'
' -----------------------------------------------------------------------------
' POURQUOI LES NOTES SONT ANALYSÉES SÉPARÉMENT
' -----------------------------------------------------------------------------
'
' Microsoft Project expose notamment les propriétés :
'
'       Task.Objects
'       Resource.Objects
'       LinkedFields
'
' Ces propriétés permettent de détecter plusieurs types d'objets ou de liens
' OLE. Elles ne permettent toutefois pas de conclure que le champ Notes est
' exempt de contenu OLE.
'
' Un objet incorporé dans une Note peut donc produire la situation suivante :
'
'       Objects = 0
'       LinkedFields = False
'
' alors qu'un avertissement OLE est quand même déclenché lorsque Project tente
' de lire ou d'afficher cette Note.
'
' La macro complète effectue donc volontairement une lecture de la propriété
' .Notes de chaque élément.
'
'
' -----------------------------------------------------------------------------
' DÉTECTION COMPORTEMENTALE DES OLE DANS LES NOTES
' -----------------------------------------------------------------------------
'
' Project ne fournit pas une propriété simple permettant de demander :
'
'       « Est-ce que cette Note contient un objet OLE? »
'
' Cependant, lorsqu'un objet OLE problématique est rencontré pendant la
' lecture de .Notes, Project peut afficher son dialogue de sécurité OLE.
'
' Pendant que ce dialogue est affiché, l'exécution VBA est suspendue.
'
' La macro mesure donc le temps nécessaire pour lire chaque Note :
'
'   1. le chronomètre démarre;
'   2. VBA tente de lire .Notes;
'   3. si Project présente le dialogue OLE, l'exécution est suspendue;
'   4. le pilote répond « Non »;
'   5. VBA reprend son traitement;
'   6. le temps écoulé est calculé;
'   7. une durée anormalement élevée marque la Note :
'
'          OLE_NOTES = Oui (suspect)
'
' Cette méthode est volontairement qualifiée de « suspect » puisqu'il s'agit
' d'une détection comportementale et non d'une propriété native de Project.
'
'
' -----------------------------------------------------------------------------
' UTILISATION - DIAGNOSTIC COMPLET
' -----------------------------------------------------------------------------
'
' 1. Ouvrir le plan à analyser dans Microsoft Project Desktop.
'
' 2. Appuyer sur :
'
'       ALT + F11
'
' 3. Si le module est distribué sous forme de fichier .bas :
'
'       File > Import File...
'
'    puis sélectionner :
'
'       Diagnostic_OLE_Project.bas
'
' 4. Revenir dans Project et appuyer sur :
'
'       ALT + F8
'
' 5. Exécuter :
'
'       DiagnosticOLE_Complet
'
' 6. Si Microsoft Project présente son avertissement de sécurité OLE :
'
'       - regarder la barre d'état de Project;
'       - elle indique l'élément actuellement analysé;
'       - attendre idéalement environ une seconde;
'       - répondre « Non ».
'
'    Le délai d'environ une seconde permet à la détection comportementale
'    d'identifier clairement l'élément dans le rapport.
'
' 7. À la fin du traitement, consulter le rapport CSV créé dans :
'
'       C:\Macro
'
'
' -----------------------------------------------------------------------------
' UTILISATION - DIAGNOSTIC RAPIDE
' -----------------------------------------------------------------------------
'
' La macro :
'
'       DiagnosticOLE_Rapide
'
' ne lit PAS les Notes.
'
' Elle permet d'effectuer un premier inventaire des objets OLE directement
' exposés par Project et des LinkedFields sans provoquer volontairement la
' lecture des Notes.
'
' Les résultats du diagnostic rapide sont affichés dans la fenêtre Immediate :
'
'       ALT + F11
'       CTRL + G
'
'
' -----------------------------------------------------------------------------
' RAPPORT CSV
' -----------------------------------------------------------------------------
'
' Le dossier suivant est créé automatiquement s'il n'existe pas :
'
'       C:\Macro
'
' Le fichier produit porte un nom semblable à :
'
'       Diagnostic_OLE_NomDuProjet_20260806_103000.csv
'
' Le fichier utilise :
'
'       Encodage   : UTF-8
'       Séparateur : ;
'
' L'encodage UTF-8 est utilisé afin de conserver correctement les caractères
' accentués dans un environnement francophone québécois.
'
'
' -----------------------------------------------------------------------------
' INTERPRÉTATION DES PRINCIPALES COLONNES DU CSV
' -----------------------------------------------------------------------------
'
' TYPE
'       Type d'élément analysé :
'       PROJET, TÂCHE, RESSOURCE ou AFFECTATION.
'
' ID
'       ID de l'élément dans Project lorsqu'il est disponible.
'
' UID
'       Unique ID Project. Il demeure utile lorsque les ID visibles peuvent
'       changer à la suite de modifications apportées au plan.
'
' TÂCHE
'       Nom de la tâche concernée.
'
' RESSOURCE
'       Nom de la ressource concernée.
'
' OLE
'       Synthèse générale du diagnostic.
'
'       Oui :
'           au moins un indicateur OLE a été détecté.
'
'       Non :
'           aucun indicateur OLE n'a été détecté par les contrôles utilisés.
'
' OLE_DIRECT
'       Indique si Project retourne un ou plusieurs objets avec Objects.
'
' NB_OBJETS
'       Nombre d'objets OLE directs retourné par Project.
'
' OLE_LIÉ
'       Indique si LinkedFields retourne True.
'
' OLE_NOTES
'       Indique si le comportement observé pendant la lecture des Notes laisse
'       croire à la présence d'un objet OLE.
'
'       Valeur possible :
'
'           Oui (suspect)
'
' TEMPS_NOTES_SEC
'       Durée nécessaire pour lire la propriété Notes.
'
'       Cette colonne est particulièrement utile pour confirmer qu'un dialogue
'       modal OLE a suspendu l'exécution de VBA.
'
' LONGUEUR_NOTES
'       Nombre de caractères de texte retournés par Project.
'
' URL_DANS_NOTES
'       Indique si le texte de la Note contient notamment :
'
'           http://
'           https://
'           file:
'           \\serveur\partage
'
' ERREUR_NOTES
'       Numéro d'erreur retourné pendant la lecture de Notes.
'
' DIAGNOSTIC
'       Résumé du ou des indicateurs observés pour l'élément.
'
'
' -----------------------------------------------------------------------------
' INTERPRÉTATION DES URL
' -----------------------------------------------------------------------------
'
' La présence d'une URL dans une Note n'indique PAS que cette URL constitue
' un objet OLE ou qu'elle est invalide.
'
' Par exemple, une URL OpenText peut être parfaitement valide pour les
' utilisateurs autorisés même si le compte du pilote effectuant le diagnostic
' n'a pas accès au système cible.
'
' Les URL sont donc inventoriées comme information de contexte uniquement.
'
'
' -----------------------------------------------------------------------------
' LIMITES DE L'OUTIL
' -----------------------------------------------------------------------------
'
' Le résultat suivant :
'
'       OLE_DIRECT = Non
'       OLE_LIÉ = Non
'
' ne garantit pas qu'aucun objet OLE n'existe dans les Notes.
'
' De la même manière :
'
'       URL_DANS_NOTES = Oui
'
' ne constitue pas une preuve de présence d'un objet OLE.
'
' La colonne OLE_NOTES repose sur le comportement observé au moment de la
' lecture de la Note et doit être considérée comme un indicateur permettant
' au pilote de localiser rapidement l'élément à investiguer.
'
'
' -----------------------------------------------------------------------------
' SÉCURITÉ
' -----------------------------------------------------------------------------
'
' Cet outil sert à identifier les objets problématiques. Il ne cherche pas à
' les activer, à les supprimer ou à modifier le contenu du plan.
'
' Lorsqu'un avertissement OLE apparaît pendant un diagnostic :
'
'       privilégier « Non »
'
' sauf si la source est connue, approuvée et que son activation est
' explicitement requise.
'
'
' -----------------------------------------------------------------------------
' MODIFICATION FUTURE DU MODULE
' -----------------------------------------------------------------------------
'
' Les fonctions de lecture sécurisée utilisent volontairement :
'
'       On Error Resume Next
'
' uniquement autour de propriétés Project particulières.
'
' Cette approche évite qu'un objet atypique interrompe complètement
' l'inventaire.
'
' Elle ne doit pas être généralisée à l'ensemble de la macro, puisqu'elle
' pourrait masquer des erreurs de programmation importantes.
'
' Le module utilise également ADODB.Stream en late binding afin de produire
' un fichier CSV UTF-8 sans nécessiter l'ajout manuel d'une référence ADODB
' dans l'éditeur VBA.
'
' =============================================================================


' =============================================================================
' CONFIGURATION
' =============================================================================

' Une lecture normale de Notes est généralement très rapide.
'
' Lorsqu'une boîte modale OLE apparaît, VBA demeure suspendu jusqu'à ce que
' le pilote réponde.
'
' Toute lecture dont la durée dépasse ce seuil est considérée comme suspecte.
'
' Le seuil est exprimé en secondes.
'
Private Const SEUIL_OLE_NOTES As Double = 0.2

' Dossier dans lequel les rapports CSV sont produits.
'
Private Const DOSSIER_RAPPORT As String = "C:\Macro"


' =============================================================================
' DIAGNOSTIC OLE COMPLET
' =============================================================================
'
' Cette macro :
'
'   - analyse la tâche récapitulative du projet;
'   - analyse toutes les tâches;
'   - analyse les affectations de chaque tâche;
'   - analyse toutes les ressources;
'   - vérifie Objects;
'   - vérifie LinkedFields;
'   - lit les Notes;
'   - mesure le temps de lecture des Notes;
'   - recherche les URL et chemins UNC;
'   - produit un rapport CSV UTF-8 dans C:\Macro.
'
' =============================================================================

Public Sub DiagnosticOLE_Complet()

    Dim p As Project
    Dim t As Task
    Dim r As Resource
    Dim a As Assignment
    Dim t0 As Task

    Dim ts As Object
    Dim chemin As String

    Dim nbObj As Long
    Dim lie As Boolean

    Dim notesTxt As String
    Dim notesSec As Double
    Dim notesErr As Long
    Dim notesSuspect As Boolean

    Dim oleGlobal As Boolean

    Dim nbSuspects As Long
    Dim nbLignes As Long


    ' -------------------------------------------------------------------------
    ' Vérifier qu'un projet est bien ouvert.
    ' -------------------------------------------------------------------------

    If ActiveProject Is Nothing Then

        MsgBox _
            "Aucun projet actif n'a été trouvé.", _
            vbExclamation, _
            "Diagnostic OLE Project"

        Exit Sub

    End If

    Set p = ActiveProject


    ' -------------------------------------------------------------------------
    ' Préparer le dossier de sortie.
    '
    ' Le dossier C:\Macro est créé automatiquement s'il n'existe pas.
    ' -------------------------------------------------------------------------

    If Not PreparerDossierRapport(DOSSIER_RAPPORT) Then

        MsgBox _
            "Impossible de créer ou d'utiliser le dossier :" & _
            vbCrLf & _
            DOSSIER_RAPPORT, _
            vbCritical, _
            "Diagnostic OLE Project"

        Exit Sub

    End If


    ' -------------------------------------------------------------------------
    ' Construire le nom du rapport.
    '
    ' L'horodatage évite d'écraser un rapport produit précédemment.
    ' -------------------------------------------------------------------------

    chemin = _
        DOSSIER_RAPPORT & _
        "\Diagnostic_OLE_" & _
        NettoyerNomFichier(p.Name) & _
        "_" & _
        Format$(Now, "yyyymmdd_hhnnss") & _
        ".csv"


    ' -------------------------------------------------------------------------
    ' Créer un flux texte UTF-8.
    '
    ' ADODB.Stream est instancié en late binding.
    ' Aucune référence ADODB ne doit donc être ajoutée manuellement dans VBA.
    ' -------------------------------------------------------------------------

    Set ts = CreerFluxUTF8()

    If ts Is Nothing Then

        MsgBox _
            "Impossible de créer le flux UTF-8 nécessaire au rapport CSV.", _
            vbCritical, _
            "Diagnostic OLE Project"

        Exit Sub

    End If


    ' -------------------------------------------------------------------------
    ' Écrire l'en-tête du rapport.
    ' -------------------------------------------------------------------------

    ts.WriteText _
        "TYPE;" & _
        "ID;" & _
        "UID;" & _
        "TÂCHE;" & _
        "RESSOURCE;" & _
        "OLE;" & _
        "OLE_DIRECT;" & _
        "NB_OBJETS;" & _
        "OLE_LIÉ;" & _
        "OLE_NOTES;" & _
        "TEMPS_NOTES_SEC;" & _
        "LONGUEUR_NOTES;" & _
        "URL_DANS_NOTES;" & _
        "ERREUR_NOTES;" & _
        "DIAGNOSTIC" & _
        vbCrLf


    ' -------------------------------------------------------------------------
    ' Initialiser le journal de diagnostic dans la fenêtre Immediate.
    ' -------------------------------------------------------------------------

    Debug.Print
    Debug.Print String(120, "=")

    Debug.Print _
        "DIAGNOSTIC OLE COMPLET - " & _
        p.Name

    Debug.Print _
        "IMPORTANT : si la boîte de sécurité OLE apparaît, cliquer sur NON."

    Debug.Print _
        "Rapport CSV : " & _
        chemin

    Debug.Print String(120, "=")


    ' =========================================================================
    ' 1. TÂCHE RÉCAPITULATIVE DU PROJET
    ' =========================================================================
    '
    ' La tâche récapitulative correspond à la tâche 0 du projet.
    '
    ' Elle peut elle aussi contenir des Notes ou certains éléments hérités
    ' d'un plan historique.
    ' =========================================================================

    On Error Resume Next

    Set t0 = p.ProjectSummaryTask

    On Error GoTo 0


    If Not t0 Is Nothing Then


        ' ---------------------------------------------------------------------
        ' Mettre à jour la barre d'état AVANT toute lecture susceptible de
        ' déclencher le dialogue OLE.
        ' ---------------------------------------------------------------------

        Application.StatusBar = _
            "Diagnostic OLE > PROJET / Tâche récapitulative"

        DoEvents


        ' ---------------------------------------------------------------------
        ' Vérification des objets OLE directs.
        ' ---------------------------------------------------------------------

        nbObj = _
            SafeTaskObjects(t0)


        ' ---------------------------------------------------------------------
        ' Vérification des champs liés par OLE.
        ' ---------------------------------------------------------------------

        lie = _
            SafeTaskLinkedFields(t0)


        ' ---------------------------------------------------------------------
        ' Lecture contrôlée des Notes.
        '
        ' Cette instruction peut déclencher le dialogue de sécurité OLE.
        ' ---------------------------------------------------------------------

        notesTxt = _
            LireTaskNotes( _
                t0, _
                notesSec, _
                notesErr)


        ' ---------------------------------------------------------------------
        ' Déterminer si le temps de lecture indique un comportement suspect.
        ' ---------------------------------------------------------------------

        notesSuspect = _
            EstNoteOLESuspecte( _
                notesSec, _
                notesErr)


        ' ---------------------------------------------------------------------
        ' Construire l'indicateur OLE global de l'élément.
        ' ---------------------------------------------------------------------

        oleGlobal = _
            (nbObj > 0 Or _
             lie Or _
             notesSuspect)


        If notesSuspect Then

            nbSuspects = _
                nbSuspects + 1


            Debug.Print _
                "*** SUSPECT OLE NOTES *** " & _
                "PROJET" & _
                " | UID=" & _
                SafeTaskUID(t0) & _
                " | " & _
                t0.Name & _
                " | temps=" & _
                Format$(notesSec, "0.000") & _
                " s"

        End If


        ' ---------------------------------------------------------------------
        ' Ajouter la tâche récapitulative au rapport CSV.
        ' ---------------------------------------------------------------------

        EcrireLigne _
            ts, _
            "PROJET", _
            "0", _
            SafeTaskUID(t0), _
            t0.Name, _
            "", _
            OuiNon(oleGlobal), _
            OuiNon(nbObj > 0), _
            CStr(nbObj), _
            OuiNon(lie), _
            StatutNotes(notesSuspect), _
            Format$(notesSec, "0.000"), _
            CStr(Len(notesTxt)), _
            OuiNon(ContientURL(notesTxt)), _
            CStr(notesErr), _
            ConstruireDiagnostic( _
                nbObj, _
                lie, _
                notesSuspect, _
                notesErr)


        nbLignes = _
            nbLignes + 1

    End If


    ' =========================================================================
    ' 2. TÂCHES ET AFFECTATIONS
    ' =========================================================================
    '
    ' Chaque tâche est analysée séparément.
    '
    ' Les affectations sont ensuite parcourues à partir de :
    '
    '       Task.Assignments
    '
    ' plutôt qu'avec ActiveProject.Assignments.
    '
    ' Cette méthode correspond au modèle objet Project et évite notamment
    ' l'erreur d'exécution 438 rencontrée pendant le développement initial
    ' de l'outil.
    ' =========================================================================

    For Each t In p.Tasks

        If Not t Is Nothing Then


            ' -----------------------------------------------------------------
            ' Afficher la tâche en cours AVANT de lire ses Notes.
            '
            ' Si le dialogue OLE apparaît, cette information permet au pilote
            ' d'identifier immédiatement l'élément en cours d'analyse.
            ' -----------------------------------------------------------------

            Application.StatusBar = _
                "Diagnostic OLE > TÂCHE ID " & _
                t.ID & _
                " / UID " & _
                t.UniqueID & _
                " > " & _
                t.Name

            DoEvents


            ' -----------------------------------------------------------------
            ' Vérifier les objets OLE directs.
            ' -----------------------------------------------------------------

            nbObj = _
                SafeTaskObjects(t)


            ' -----------------------------------------------------------------
            ' Vérifier les champs liés par OLE.
            ' -----------------------------------------------------------------

            lie = _
                SafeTaskLinkedFields(t)


            ' -----------------------------------------------------------------
            ' Lire volontairement les Notes.
            '
            ' Cette lecture est nécessaire pour détecter les objets qui ne sont
            ' pas visibles avec Objects ou LinkedFields.
            ' -----------------------------------------------------------------

            notesTxt = _
                LireTaskNotes( _
                    t, _
                    notesSec, _
                    notesErr)


            notesSuspect = _
                EstNoteOLESuspecte( _
                    notesSec, _
                    notesErr)


            oleGlobal = _
                (nbObj > 0 Or _
                 lie Or _
                 notesSuspect)


            If notesSuspect Then

                nbSuspects = _
                    nbSuspects + 1


                Debug.Print _
                    "*** SUSPECT OLE NOTES *** " & _
                    "TÂCHE" & _
                    " | ID=" & _
                    t.ID & _
                    " | UID=" & _
                    t.UniqueID & _
                    " | " & _
                    t.Name & _
                    " | temps=" & _
                    Format$(notesSec, "0.000") & _
                    " s"

            End If


            ' -----------------------------------------------------------------
            ' Ajouter la tâche au rapport CSV.
            ' -----------------------------------------------------------------

            EcrireLigne _
                ts, _
                "TÂCHE", _
                CStr(t.ID), _
                CStr(t.UniqueID), _
                t.Name, _
                "", _
                OuiNon(oleGlobal), _
                OuiNon(nbObj > 0), _
                CStr(nbObj), _
                OuiNon(lie), _
                StatutNotes(notesSuspect), _
                Format$(notesSec, "0.000"), _
                CStr(Len(notesTxt)), _
                OuiNon(ContientURL(notesTxt)), _
                CStr(notesErr), _
                ConstruireDiagnostic( _
                    nbObj, _
                    lie, _
                    notesSuspect, _
                    notesErr)


            nbLignes = _
                nbLignes + 1


            ' =================================================================
            ' 2.1 AFFECTATIONS DE LA TÂCHE
            ' =================================================================
            '
            ' Une affectation représente l'association entre une tâche et une
            ' ressource.
            '
            ' Les affectations sont parcourues depuis la tâche actuelle.
            ' =================================================================

            For Each a In t.Assignments


                Application.StatusBar = _
                    "Diagnostic OLE > AFFECTATION > Tâche ID " & _
                    t.ID & _
                    " > " & _
                    t.Name & _
                    " > Ressource : " & _
                    SafeAssignmentResourceName(a)

                DoEvents


                ' -------------------------------------------------------------
                ' Les affectations ne disposent pas du même compteur Objects
                ' que les tâches et ressources.
                '
                ' Le diagnostic vérifie donc LinkedFields et les Notes.
                ' -------------------------------------------------------------

                lie = _
                    SafeAssignmentLinkedFields(a)


                notesTxt = _
                    LireAssignmentNotes( _
                        a, _
                        notesSec, _
                        notesErr)


                notesSuspect = _
                    EstNoteOLESuspecte( _
                        notesSec, _
                        notesErr)


                oleGlobal = _
                    (lie Or _
                     notesSuspect)


                If notesSuspect Then

                    nbSuspects = _
                        nbSuspects + 1


                    Debug.Print _
                        "*** SUSPECT OLE NOTES *** " & _
                        "AFFECTATION" & _
                        " | UID=" & _
                        SafeAssignmentUID(a) & _
                        " | Tâche ID=" & _
                        t.ID & _
                        " | " & _
                        t.Name & _
                        " | Ressource=" & _
                        SafeAssignmentResourceName(a) & _
                        " | temps=" & _
                        Format$(notesSec, "0.000") & _
                        " s"

                End If


                ' -------------------------------------------------------------
                ' Ajouter l'affectation au rapport CSV.
                '
                ' OLE_DIRECT est indiqué S.O. puisque le contrôle Objects n'est
                ' pas utilisé pour les affectations.
                ' -------------------------------------------------------------

                EcrireLigne _
                    ts, _
                    "AFFECTATION", _
                    "", _
                    SafeAssignmentUID(a), _
                    t.Name, _
                    SafeAssignmentResourceName(a), _
                    OuiNon(oleGlobal), _
                    "S.O.", _
                    "", _
                    OuiNon(lie), _
                    StatutNotes(notesSuspect), _
                    Format$(notesSec, "0.000"), _
                    CStr(Len(notesTxt)), _
                    OuiNon(ContientURL(notesTxt)), _
                    CStr(notesErr), _
                    ConstruireDiagnostic( _
                        0, _
                        lie, _
                        notesSuspect, _
                        notesErr)


                nbLignes = _
                    nbLignes + 1


            Next a

        End If

    Next t


    ' =========================================================================
    ' 3. RESSOURCES
    ' =========================================================================
    '
    ' Les ressources sont analysées après les tâches et affectations.
    '
    ' Comme pour les tâches, trois contrôles sont effectués :
    '
    '       Objects
    '       LinkedFields
    '       Notes
    ' =========================================================================

    For Each r In p.Resources

        If Not r Is Nothing Then


            Application.StatusBar = _
                "Diagnostic OLE > RESSOURCE ID " & _
                r.ID & _
                " / UID " & _
                r.UniqueID & _
                " > " & _
                r.Name

            DoEvents


            ' -----------------------------------------------------------------
            ' Vérifier les objets OLE directs.
            ' -----------------------------------------------------------------

            nbObj = _
                SafeResourceObjects(r)


            ' -----------------------------------------------------------------
            ' Vérifier les champs liés par OLE.
            ' -----------------------------------------------------------------

            lie = _
                SafeResourceLinkedFields(r)


            ' -----------------------------------------------------------------
            ' Lire les Notes de la ressource.
            ' -----------------------------------------------------------------

            notesTxt = _
                LireResourceNotes( _
                    r, _
                    notesSec, _
                    notesErr)


            notesSuspect = _
                EstNoteOLESuspecte( _
                    notesSec, _
                    notesErr)


            oleGlobal = _
                (nbObj > 0 Or _
                 lie Or _
                 notesSuspect)


            If notesSuspect Then

                nbSuspects = _
                    nbSuspects + 1


                Debug.Print _
                    "*** SUSPECT OLE NOTES *** " & _
                    "RESSOURCE" & _
                    " | ID=" & _
                    r.ID & _
                    " | UID=" & _
                    r.UniqueID & _
                    " | " & _
                    r.Name & _
                    " | temps=" & _
                    Format$(notesSec, "0.000") & _
                    " s"

            End If


            ' -----------------------------------------------------------------
            ' Ajouter la ressource au rapport CSV.
            ' -----------------------------------------------------------------

            EcrireLigne _
                ts, _
                "RESSOURCE", _
                CStr(r.ID), _
                CStr(r.UniqueID), _
                "", _
                r.Name, _
                OuiNon(oleGlobal), _
                OuiNon(nbObj > 0), _
                CStr(nbObj), _
                OuiNon(lie), _
                StatutNotes(notesSuspect), _
                Format$(notesSec, "0.000"), _
                CStr(Len(notesTxt)), _
                OuiNon(ContientURL(notesTxt)), _
                CStr(notesErr), _
                ConstruireDiagnostic( _
                    nbObj, _
                    lie, _
                    notesSuspect, _
                    notesErr)


            nbLignes = _
                nbLignes + 1

        End If

    Next r


    ' =========================================================================
    ' 4. SAUVEGARDE DU RAPPORT
    ' =========================================================================

    Application.StatusBar = _
        "Diagnostic OLE > Enregistrement du rapport..."

    DoEvents


    On Error GoTo ErreurSauvegarde


    ' ADODB.Stream conserve le contenu en mémoire jusqu'à SaveToFile.
    '
    ' Le paramètre 2 correspond à adSaveCreateOverWrite.
    '
    ts.SaveToFile chemin, 2

    ts.Close

    Set ts = Nothing


    Application.StatusBar = ""


    ' =========================================================================
    ' 5. SOMMAIRE
    ' =========================================================================

    Debug.Print String(120, "=")

    Debug.Print _
        "FIN - lignes inventoriées : " & _
        nbLignes

    Debug.Print _
        "Suspects OLE dans Notes   : " & _
        nbSuspects

    Debug.Print _
        "CSV : " & _
        chemin

    Debug.Print String(120, "=")


    MsgBox _
        "Diagnostic OLE terminé." & _
        vbCrLf & vbCrLf & _
        "Lignes inventoriées : " & _
        nbLignes & _
        vbCrLf & _
        "Suspects OLE dans Notes : " & _
        nbSuspects & _
        vbCrLf & vbCrLf & _
        "Rapport :" & _
        vbCrLf & _
        chemin & _
        vbCrLf & vbCrLf & _
        "Dans Excel, filtrer d'abord la colonne OLE sur Oui, " & _
        "puis vérifier OLE_NOTES = Oui (suspect).", _
        vbInformation, _
        "Diagnostic OLE Project"


    Exit Sub


' =============================================================================
' GESTION D'UNE ERREUR DE SAUVEGARDE
' =============================================================================

ErreurSauvegarde:

    Application.StatusBar = ""

    On Error Resume Next

    If Not ts Is Nothing Then
        ts.Close
    End If

    Set ts = Nothing

    On Error GoTo 0


    MsgBox _
        "Le rapport CSV n'a pas pu être enregistré." & _
        vbCrLf & vbCrLf & _
        "Chemin :" & _
        vbCrLf & _
        chemin & _
        vbCrLf & vbCrLf & _
        "Erreur : " & _
        Err.Number & _
        vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Diagnostic OLE Project"

End Sub


' =============================================================================
' DIAGNOSTIC OLE RAPIDE
' =============================================================================
'
' Ce scan ne lit PAS les Notes.
'
' Il permet d'obtenir rapidement :
'
'   - les objets OLE directs des tâches;
'   - les objets OLE directs des ressources;
'   - les LinkedFields des tâches;
'   - les LinkedFields des ressources;
'   - les LinkedFields des affectations.
'
' Cette macro est utile lorsqu'on veut effectuer un premier inventaire sans
' volontairement déclencher les avertissements OLE associés aux Notes.
'
' Les résultats sont écrits dans la fenêtre Immediate :
'
'       ALT + F11
'       CTRL + G
'
' =============================================================================

Public Sub DiagnosticOLE_Rapide()

    Dim p As Project
    Dim t As Task
    Dim r As Resource
    Dim a As Assignment


    If ActiveProject Is Nothing Then

        MsgBox _
            "Aucun projet actif n'a été trouvé.", _
            vbExclamation, _
            "Diagnostic OLE Project"

        Exit Sub

    End If


    Set p = ActiveProject


    Debug.Print
    Debug.Print String(100, "=")

    Debug.Print _
        "DIAGNOSTIC OLE RAPIDE - " & _
        p.Name

    Debug.Print _
        "TYPE", _
        "ID", _
        "UID", _
        "OLE", _
        "OLE DIRECT", _
        "OLE LIÉ", _
        "NOM"

    Debug.Print String(100, "-")


    ' -------------------------------------------------------------------------
    ' Tâches et affectations.
    ' -------------------------------------------------------------------------

    For Each t In p.Tasks

        If Not t Is Nothing Then


            Debug.Print _
                "TÂCHE", _
                t.ID, _
                t.UniqueID, _
                OuiNon( _
                    SafeTaskObjects(t) > 0 Or _
                    SafeTaskLinkedFields(t)), _
                OuiNon( _
                    SafeTaskObjects(t) > 0), _
                OuiNon( _
                    SafeTaskLinkedFields(t)), _
                t.Name


            ' Les affectations sont parcourues depuis la tâche.
            '
            For Each a In t.Assignments


                Debug.Print _
                    "AFFECTATION", _
                    "", _
                    SafeAssignmentUID(a), _
                    OuiNon( _
                        SafeAssignmentLinkedFields(a)), _
                    "S.O.", _
                    OuiNon( _
                        SafeAssignmentLinkedFields(a)), _
                    t.Name & _
                    " / " & _
                    SafeAssignmentResourceName(a)


            Next a

        End If

    Next t


    ' -------------------------------------------------------------------------
    ' Ressources.
    ' -------------------------------------------------------------------------

    For Each r In p.Resources

        If Not r Is Nothing Then


            Debug.Print _
                "RESSOURCE", _
                r.ID, _
                r.UniqueID, _
                OuiNon( _
                    SafeResourceObjects(r) > 0 Or _
                    SafeResourceLinkedFields(r)), _
                OuiNon( _
                    SafeResourceObjects(r) > 0), _
                OuiNon( _
                    SafeResourceLinkedFields(r)), _
                r.Name


        End If

    Next r


    Debug.Print String(100, "=")


    MsgBox _
        "Diagnostic rapide terminé." & _
        vbCrLf & vbCrLf & _
        "Consulter la fenêtre Immediate avec CTRL + G.", _
        vbInformation, _
        "Diagnostic OLE Project"

End Sub


' =============================================================================
' LECTURE CONTRÔLÉE DES NOTES - TÂCHES
' =============================================================================
'
' Cette fonction lit le champ Notes d'une tâche tout en mesurant la durée de
' l'opération.
'
' Si Project présente son dialogue de sécurité OLE, l'exécution VBA est
' suspendue. Cette suspension se reflète dans la valeur retournée dans duree.
'
' numeroErreur permet de conserver une éventuelle erreur de lecture dans
' le rapport sans interrompre l'ensemble du diagnostic.
'
' =============================================================================

Private Function LireTaskNotes( _
    ByVal t As Task, _
    ByRef duree As Double, _
    ByRef numeroErreur As Long) As String


    Dim debut As Single
    Dim fin As Single


    debut = Timer


    On Error Resume Next

    Err.Clear


    LireTaskNotes = _
        t.Notes


    numeroErreur = _
        Err.Number


    Err.Clear

    On Error GoTo 0


    fin = Timer


    duree = _
        DureeEcoulee( _
            debut, _
            fin)

End Function


' =============================================================================
' LECTURE CONTRÔLÉE DES NOTES - RESSOURCES
' =============================================================================
'
' Même principe que LireTaskNotes, mais appliqué aux ressources.
'
' =============================================================================

Private Function LireResourceNotes( _
    ByVal r As Resource, _
    ByRef duree As Double, _
    ByRef numeroErreur As Long) As String


    Dim debut As Single
    Dim fin As Single


    debut = Timer


    On Error Resume Next

    Err.Clear


    LireResourceNotes = _
        r.Notes


    numeroErreur = _
        Err.Number


    Err.Clear

    On Error GoTo 0


    fin = Timer


    duree = _
        DureeEcoulee( _
            debut, _
            fin)

End Function


' =============================================================================
' LECTURE CONTRÔLÉE DES NOTES - AFFECTATIONS
' =============================================================================
'
' Même principe que LireTaskNotes, mais appliqué aux affectations.
'
' =============================================================================

Private Function LireAssignmentNotes( _
    ByVal a As Assignment, _
    ByRef duree As Double, _
    ByRef numeroErreur As Long) As String


    Dim debut As Single
    Dim fin As Single


    debut = Timer


    On Error Resume Next

    Err.Clear


    LireAssignmentNotes = _
        a.Notes


    numeroErreur = _
        Err.Number


    Err.Clear

    On Error GoTo 0


    fin = Timer


    duree = _
        DureeEcoulee( _
            debut, _
            fin)

End Function


' =============================================================================
' CALCUL DU TEMPS ÉCOULÉ
' =============================================================================
'
' Timer retourne le nombre de secondes écoulées depuis minuit.
'
' Le calcul gère également le cas exceptionnel où le traitement traverserait
' minuit pendant une lecture.
'
' =============================================================================

Private Function DureeEcoulee( _
    ByVal debut As Single, _
    ByVal fin As Single) As Double


    If fin >= debut Then

        DureeEcoulee = _
            CDbl(fin - debut)

    Else

        ' Passage de minuit.
        '
        DureeEcoulee = _
            CDbl( _
                (86400! - debut) + _
                fin)

    End If

End Function


' =============================================================================
' DÉTERMINER SI UNE NOTE EST SUSPECTE
' =============================================================================
'
' numeroErreur est fourni à la fonction pour conserver la signature et le
' contexte du diagnostic, mais le numéro d'erreur n'est PAS utilisé seul pour
' conclure qu'un objet OLE est présent.
'
' Le principal signal utilisé est le délai observé lors de la lecture.
'
' =============================================================================

Private Function EstNoteOLESuspecte( _
    ByVal duree As Double, _
    ByVal numeroErreur As Long) As Boolean


    EstNoteOLESuspecte = _
        (duree >= SEUIL_OLE_NOTES)

End Function


' =============================================================================
' LECTURE SÉCURISÉE DE TASK.OBJECTS
' =============================================================================
'
' Certaines propriétés Project peuvent retourner une erreur selon le contenu
' ou l'état de l'objet.
'
' On Error Resume Next est limité volontairement à cette lecture précise.
'
' Une erreur est interprétée ici comme :
'
'       nombre d'objets inconnu = 0
'
' afin de permettre à l'inventaire de continuer.
'
' =============================================================================

Private Function SafeTaskObjects( _
    ByVal t As Task) As Long


    On Error Resume Next

    Err.Clear


    SafeTaskObjects = _
        t.Objects


    If Err.Number <> 0 Then

        SafeTaskObjects = 0

    End If


    Err.Clear

    On Error GoTo 0

End Function


' =============================================================================
' LECTURE SÉCURISÉE DE RESOURCE.OBJECTS
' =============================================================================

Private Function SafeResourceObjects( _
    ByVal r As Resource) As Long


    On Error Resume Next

    Err.Clear


    SafeResourceObjects = _
        r.Objects


    If Err.Number <> 0 Then

        SafeResourceObjects = 0

    End If


    Err.Clear

    On Error GoTo 0

End Function


' =============================================================================
' LECTURE SÉCURISÉE DE TASK.LINKEDFIELDS
' =============================================================================
'
' LinkedFields indique si un champ de la tâche est lié à une autre
' application au moyen d'OLE.
'
' =============================================================================

Private Function SafeTaskLinkedFields( _
    ByVal t As Task) As Boolean


    On Error Resume Next

    Err.Clear


    SafeTaskLinkedFields = _
        t.LinkedFields


    If Err.Number <> 0 Then

        SafeTaskLinkedFields = False

    End If


    Err.Clear

    On Error GoTo 0

End Function


' =============================================================================
' LECTURE SÉCURISÉE DE RESOURCE.LINKEDFIELDS
' =============================================================================

Private Function SafeResourceLinkedFields( _
    ByVal r As Resource) As Boolean


    On Error Resume Next

    Err.Clear


    SafeResourceLinkedFields = _
        r.LinkedFields


    If Err.Number <> 0 Then

        SafeResourceLinkedFields = False

    End If


    Err.Clear

    On Error GoTo 0

End Function


' =============================================================================
' LECTURE SÉCURISÉE DE ASSIGNMENT.LINKEDFIELDS
' =============================================================================

Private Function SafeAssignmentLinkedFields( _
    ByVal a As Assignment) As Boolean


    On Error Resume Next

    Err.Clear


    SafeAssignmentLinkedFields = _
        a.LinkedFields


    If Err.Number <> 0 Then

        SafeAssignmentLinkedFields = False

    End If


    Err.Clear

    On Error GoTo 0

End Function


' =============================================================================
' RÉCUPÉRATION SÉCURISÉE DU UID D'UNE TÂCHE
' =============================================================================

Private Function SafeTaskUID( _
    ByVal t As Task) As String


    On Error Resume Next

    Err.Clear


    SafeTaskUID = _
        CStr(t.UniqueID)


    If Err.Number <> 0 Then

        SafeTaskUID = ""

    End If


    Err.Clear

    On Error GoTo 0

End Function


' =============================================================================
' RÉCUPÉRATION SÉCURISÉE DU UID D'UNE AFFECTATION
' =============================================================================

Private Function SafeAssignmentUID( _
    ByVal a As Assignment) As String


    On Error Resume Next

    Err.Clear


    SafeAssignmentUID = _
        CStr(a.UniqueID)


    If Err.Number <> 0 Then

        SafeAssignmentUID = ""

    End If


    Err.Clear

    On Error GoTo 0

End Function


' =============================================================================
' RÉCUPÉRATION SÉCURISÉE DU NOM DE RESSOURCE D'UNE AFFECTATION
' =============================================================================

Private Function SafeAssignmentResourceName( _
    ByVal a As Assignment) As String


    On Error Resume Next

    Err.Clear


    SafeAssignmentResourceName = _
        a.ResourceName


    If Err.Number <> 0 Then

        SafeAssignmentResourceName = ""

    End If


    Err.Clear

    On Error GoTo 0

End Function


' =============================================================================
' PRÉPARATION DU DOSSIER DE RAPPORT
' =============================================================================
'
' Crée C:\Macro s'il n'existe pas.
'
' La fonction retourne True lorsque le dossier est disponible à la fin du
' traitement.
'
' =============================================================================

Private Function PreparerDossierRapport( _
    ByVal cheminDossier As String) As Boolean


    On Error Resume Next

    Err.Clear


    If Dir$(cheminDossier, vbDirectory) = "" Then

        MkDir cheminDossier

    End If


    PreparerDossierRapport = _
        (Err.Number = 0 And _
         Dir$(cheminDossier, vbDirectory) <> "")


    Err.Clear

    On Error GoTo 0

End Function


' =============================================================================
' CRÉATION D'UN FLUX TEXTE UTF-8
' =============================================================================
'
' FileSystemObject.CreateTextFile produit habituellement du texte ANSI ou
' Unicode UTF-16 selon ses paramètres.
'
' Pour garantir explicitement un CSV UTF-8, le module utilise ADODB.Stream.
'
' L'objet est créé en late binding :
'
'       CreateObject("ADODB.Stream")
'
' Il n'est donc pas nécessaire d'activer une référence supplémentaire dans :
'
'       Tools > References
'
' =============================================================================

Private Function CreerFluxUTF8() As Object


    Dim flux As Object


    On Error Resume Next

    Err.Clear


    Set flux = _
        CreateObject("ADODB.Stream")


    If Err.Number <> 0 Or flux Is Nothing Then

        Set CreerFluxUTF8 = Nothing

        Err.Clear

        On Error GoTo 0

        Exit Function

    End If


    ' 2 = adTypeText
    '
    flux.Type = 2


    ' Forcer explicitement l'encodage UTF-8.
    '
    flux.Charset = "utf-8"


    flux.Open


    Set CreerFluxUTF8 = _
        flux


    On Error GoTo 0

End Function


' =============================================================================
' ÉCRITURE D'UNE LIGNE DANS LE RAPPORT CSV
' =============================================================================
'
' Le séparateur utilisé est volontairement le point-virgule :
'
'       ;
'
' ce qui correspond au comportement attendu dans Excel dans un environnement
' francophone utilisant la virgule comme séparateur décimal.
'
' Toutes les valeurs sont entourées de guillemets afin de protéger les noms,
' Notes ou autres champs pouvant contenir le séparateur.
'
' =============================================================================

Private Sub EcrireLigne( _
    ByVal ts As Object, _
    ByVal typeElement As String, _
    ByVal ID As String, _
    ByVal UID As String, _
    ByVal nomTache As String, _
    ByVal nomRessource As String, _
    ByVal ole As String, _
    ByVal oleDirect As String, _
    ByVal nbObjets As String, _
    ByVal oleLie As String, _
    ByVal oleNotes As String, _
    ByVal tempsNotes As String, _
    ByVal longueurNotes As String, _
    ByVal urlNotes As String, _
    ByVal erreurNotes As String, _
    ByVal diagnostic As String)


    ts.WriteText _
        CSV(typeElement) & ";" & _
        CSV(ID) & ";" & _
        CSV(UID) & ";" & _
        CSV(nomTache) & ";" & _
        CSV(nomRessource) & ";" & _
        CSV(ole) & ";" & _
        CSV(oleDirect) & ";" & _
        CSV(nbObjets) & ";" & _
        CSV(oleLie) & ";" & _
        CSV(oleNotes) & ";" & _
        CSV(tempsNotes) & ";" & _
        CSV(longueurNotes) & ";" & _
        CSV(urlNotes) & ";" & _
        CSV(erreurNotes) & ";" & _
        CSV(diagnostic) & _
        vbCrLf

End Sub


' =============================================================================
' NORMALISATION D'UNE VALEUR POUR LE CSV
' =============================================================================
'
' Les guillemets contenus dans les données sont doublés conformément au format
' CSV.
'
' Les retours de ligne sont remplacés par des espaces afin qu'une Note ne
' provoque pas artificiellement plusieurs lignes dans le rapport.
'
' =============================================================================

Private Function CSV( _
    ByVal valeur As String) As String


    valeur = _
        Replace( _
            valeur, _
            """", _
            """""")


    valeur = _
        Replace( _
            valeur, _
            vbCr, _
            " ")


    valeur = _
        Replace( _
            valeur, _
            vbLf, _
            " ")


    CSV = _
        """" & _
        valeur & _
        """"

End Function


' =============================================================================
' CONVERSION BOOLÉENNE EN OUI / NON
' =============================================================================

Private Function OuiNon( _
    ByVal valeur As Boolean) As String


    If valeur Then

        OuiNon = "Oui"

    Else

        OuiNon = "Non"

    End If

End Function


' =============================================================================
' TEXTE ASSOCIÉ À LA DÉTECTION DANS LES NOTES
' =============================================================================

Private Function StatutNotes( _
    ByVal suspect As Boolean) As String


    If suspect Then

        StatutNotes = _
            "Oui (suspect)"

    Else

        StatutNotes = _
            "Non"

    End If

End Function


' =============================================================================
' CONSTRUCTION DU MESSAGE DE DIAGNOSTIC
' =============================================================================
'
' Plusieurs signaux peuvent être présents simultanément.
'
' Exemple :
'
'       Objet OLE direct détecté | Champ lié par OLE
'
' =============================================================================

Private Function ConstruireDiagnostic( _
    ByVal nbObjets As Long, _
    ByVal lie As Boolean, _
    ByVal notesSuspect As Boolean, _
    ByVal notesErr As Long) As String


    Dim s As String


    If nbObjets > 0 Then

        s = _
            AjouterTexte( _
                s, _
                "Objet OLE direct détecté")

    End If


    If lie Then

        s = _
            AjouterTexte( _
                s, _
                "Champ lié par OLE")

    End If


    If notesSuspect Then

        s = _
            AjouterTexte( _
                s, _
                "OLE probable dans les Notes : " & _
                "délai anormal observé pendant la lecture")

    End If


    If notesErr <> 0 Then

        s = _
            AjouterTexte( _
                s, _
                "Erreur de lecture des Notes=" & _
                CStr(notesErr))

    End If


    If s = "" Then

        s = _
            "Aucun OLE détecté par les contrôles utilisés"

    End If


    ConstruireDiagnostic = _
        s

End Function


' =============================================================================
' AJOUTER UN ÉLÉMENT AU MESSAGE DE DIAGNOSTIC
' =============================================================================

Private Function AjouterTexte( _
    ByVal base As String, _
    ByVal ajout As String) As String


    If base = "" Then

        AjouterTexte = _
            ajout

    Else

        AjouterTexte = _
            base & _
            " | " & _
            ajout

    End If

End Function


' =============================================================================
' DÉTECTION D'URL OU DE CHEMIN DANS LES NOTES
' =============================================================================
'
' Cette fonction détecte uniquement la présence textuelle de références
' typiques :
'
'       http://
'       https://
'       file:
'       \\serveur\partage
'
' Elle ne valide PAS :
'
'   - que l'URL existe;
'   - que l'utilisateur y a accès;
'   - que l'URL constitue un objet OLE;
'   - que la cible du lien est fonctionnelle.
'
' =============================================================================

Private Function ContientURL( _
    ByVal texte As String) As Boolean


    ContientURL = _
        (InStr( _
            1, _
            texte, _
            "http://", _
            vbTextCompare) > 0) _
        Or _
        (InStr( _
            1, _
            texte, _
            "https://", _
            vbTextCompare) > 0) _
        Or _
        (InStr( _
            1, _
            texte, _
            "file:", _
            vbTextCompare) > 0) _
        Or _
        (InStr( _
            1, _
            texte, _
            "\\", _
            vbTextCompare) > 0)

End Function


' =============================================================================
' NETTOYAGE DU NOM DE FICHIER
' =============================================================================
'
' Le nom du projet est intégré au nom du CSV.
'
' Les caractères interdits dans les noms de fichiers Windows sont donc
' remplacés par un trait de soulignement.
'
' =============================================================================

Private Function NettoyerNomFichier( _
    ByVal nom As String) As String


    Dim invalides As Variant
    Dim c As Variant


    invalides = _
        Array( _
            "\", _
            "/", _
            ":", _
            "*", _
            "?", _
            """", _
            "<", _
            ">", _
            "|")


    For Each c In invalides

        nom = _
            Replace( _
                nom, _
                CStr(c), _
                "_")

    Next c


    NettoyerNomFichier = _
        nom

End Function

