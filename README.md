# Calcul Ventes AE

Ce projet génère automatiquement un fichier Excel nommé **Calcul_Ventes_AE.xlsx** destiné à aider une auto-entreprise spécialisée dans la vente de matériel informatique et la prestation de services (main d'œuvre). Le tableau généré calcule :

- Le **prix fournisseur total** (prix unitaire × quantité)
- La **commission** (en fonction d’un pourcentage configurable)
- Le **prix client final** (prix fournisseur + commission)
- Le **total** incluant également la main d'œuvre

Le fichier est conçu avec une mise en forme soignée (couleurs, bordures, alternance de couleurs) pour faciliter la saisie des données et la lecture des résultats.

---

## Tutoriel d'exécution sous Windows

Suivez ces étapes pour exécuter le script et générer le fichier Excel :

### 1. Prérequis

- **Installer Python**  
  Téléchargez et installez Python depuis le site officiel :  
  [https://www.python.org/downloads/](https://www.python.org/downloads/)  
  *Assurez-vous de cocher l’option "Add Python to PATH" lors de l’installation.*

- **Vérifier pip**  
  Ouvrez l'invite de commandes (cmd) et tapez :
  ```
  pip --version
  ```
  Pip est normalement installé avec Python.

- **Installer openpyxl**  
  Dans l'invite de commandes, tapez :
  ```
  pip install openpyxl
  ```

### 2. Télécharger le projet

- **Clonez** le dépôt GitHub ou téléchargez-le en ZIP et extrayez-le dans un dossier sur votre ordinateur.

### 3. Exécuter le script

1. **Ouvrir l'invite de commandes**  
   - Appuyez sur `Win + R`, tapez `cmd` et appuyez sur Entrée.  
   - Vous pouvez également rechercher "Invite de commandes" dans le menu Démarrer.

2. **Naviguer jusqu'au dossier du projet**  
   Utilisez la commande `cd` pour accéder au dossier où se trouve le script.  
   Par exemple :
   ```
   cd C:\Users\VotreNom\Documents\Calcul_Ventes_AE
   ```

3. **Lancer le script**  
   Exécutez la commande suivante :
   ```
   python nom_du_script.py
   ```
   Remplacez `nom_du_script.py` par le nom du fichier contenant le code (par exemple, `calcul_ventes_ae.py`).

4. **Vérifier le résultat**  
   Le script génère le fichier **Calcul_Ventes_AE.xlsx** dans le même dossier. Ouvrez-le avec Excel pour vérifier que le tableau s'affiche correctement.

### 4. Résolution de problèmes

- **Python ou openpyxl non installés :**  
  Vérifiez vos installations et réinstallez si nécessaire.

- **Message d'erreur lors de l'exécution :**  
  Assurez-vous que vous exécutez le script avec la bonne version de Python et que le dossier du projet ne contient pas de fichiers conflictuels (comme un fichier nommé `openpyxl.py`).

---

## Contribution

Les contributions sont les bienvenues ! Pour proposer des modifications :

1. Créez une branche pour vos changements.
2. Soumettez une pull request avec une description détaillée de vos modifications.
