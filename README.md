# Move-PhotoByDate

## Description

`Move-PhotoByDate` est un script PowerShell qui permet de trier et d'organiser automatiquement des photos en les déplaçant vers des dossiers classés par date. Il crée des dossiers basés sur l'année et le mois de prise de vue, ou, si cette information n'est pas disponible, il utilise la date de dernière modification du fichier.

### Fonctionnalités

1. **Analyse du répertoire source** : Le script parcourt le répertoire spécifié à la recherche des fichiers photo.
2. **Récupération de la date de prise de vue** : Il extrait la date de prise de vue des photos à partir de leurs métadonnées. Si cette information n'est pas disponible, la date de dernière modification est utilisée.
3. **Création de répertoires** : Des dossiers sont créés en fonction de l'année et du mois si ces derniers n'existent pas encore dans le répertoire cible.
4. **Déplacement des photos** : Les fichiers sont déplacés vers le dossier correspondant à leur date.
5. **Affichage de la progression** : Le script affiche la progression du tri et du déplacement des fichiers en temps réel.
6. (Optionnel) **Suppression des dossiers vides** : Le script peut également supprimer les dossiers vides restants après le déplacement des photos.

## Utilisation

- Spécifiez le chemin du répertoire source contenant les photos à trier (`$PicturePath`).
- Indiquez le chemin du répertoire cible où les photos seront déplacées et organisées (`$TargetPath`).

Exécutez le script, et il se chargera de classer vos photos par année et mois dans les dossiers correspondants.

### Exemple de commande

```powershell
Move-Pictures -PicturePath "\\NAS\Photos\John a trier" -TargetPath "\\NAS\Photos\"
