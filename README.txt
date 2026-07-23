GESTION DES ÉQUIPES — Application hors ligne
================================================

QUE FAIT CETTE APPLICATION ?
-----------------------------
Elle remplace le classeur Excel "Equipes.xlsm" par une vraie application de
bureau qui fonctionne 100% hors ligne (aucune connexion internet requise
UNE FOIS installée) :

  • Onglet "Employés"        : ajouter, modifier et supprimer des employés
  • Onglets Lundi à Vendredi : horaire complet (TRC-1/2/3, quarts JOUR/SOIR,
    Opérateur / Aide opérateur / Journalier, Emballeur, Autres postes,
    2 postes Shipping, et une section "Postes divers" extensible pour
    les employés saisonniers) — chaque assignation a sa propre case
    "heures travaillées"
  • Panneau latéral sur chaque jour : cumul et heures restantes par
    employé, mis à jour automatiquement
  • Onglet "Heures (semaine)" : heures cible vs cumulées, écart
  • Export en Excel (.xlsx) en un clic depuis le menu Fichier

Toutes les données sont sauvegardées automatiquement (à chaque modification)
sur votre ordinateur — rien n'est perdu en fermant le programme, et rien
n'est jamais envoyé sur internet.


COMMENT OBTENIR LE .EXE — UN SEUL DOUBLE-CLIC
-------------------------------------------------
Je travaille dans un environnement Linux qui n'a pas accès à python.org,
donc je ne peux pas vous livrer directement un .exe déjà construit depuis
ici. La solution : "build_exe.bat" fait maintenant TOUT tout seul.

  1. Copiez tout le dossier "GestionEquipes" sur l'ordinateur Windows où
     l'application doit être utilisée (celui-ci doit avoir internet, ne
     serait-ce que pour cette étape unique).

  2. Double-cliquez sur "build_exe.bat".
     C'est tout. Vous n'avez RIEN d'autre à faire :
       - si Python n'est pas déjà installé, le script le télécharge et
         l'installe lui-même, en silence, sans aucun assistant à l'écran
       - il installe ensuite les outils nécessaires
       - puis il construit "GestionEquipes.exe"
     Cela prend 2 à 5 minutes. Une fenêtre noire affiche la progression.

  3. Une fois "TERMINÉ !" affiché, vous trouverez "GestionEquipes.exe"
     dans le même dossier. C'est votre application, prête à l'emploi :
       - déplacez-la sur le Bureau
       - créez un raccourci, épinglez-la à la barre des tâches
     Elle n'a plus jamais besoin d'internet après cette étape, même si
     Python a été installé au passage.

Cette étape n'est à faire qu'UNE SEULE FOIS, sur l'ordinateur qui servira
à construire l'exe (idéalement celui qui l'utilisera). Vous pouvez ensuite
copier "GestionEquipes.exe" seul sur d'autres postes Windows si besoin —
il fonctionne de façon autonome, sans réinstaller Python sur ces postes-là.

Remarque honnête : je n'ai pas pu tester ce script sur une vraie machine
Windows (mon environnement de travail est Linux, sans Windows disponible).
Il utilise cependant uniquement les options officielles et documentées de
l'installateur Python et de PyInstaller. Si un message d'erreur apparaît,
copiez-le-moi et je corrigerai le script.


OÙ SONT SAUVEGARDÉES MES DONNÉES ?
-------------------------------------
Dans : Documents\GestionEquipes\data.json (créé automatiquement).
Pour faire une copie de sécurité, il suffit de copier ce fichier.
Pour "réinitialiser" l'application, renommez ou supprimez ce fichier.


BESOIN D'AIDE ?
-----------------
Si "build_exe.bat" affiche une erreur, copiez le message affiché et je
pourrai vous aider à le corriger.
