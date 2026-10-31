# Validation candidate V4.1 — Excel, support WhatsApp, Smart App Control

> Statut : **candidate non publiée**. La V4.1 a besoin d'essais Windows interactifs. Ne pas modifier `updates/latest.json` avant qualification et signature de code.

## Incident corrigé conceptuellement

Une licence détectée comme `ACTIVE` via les interfaces Windows/Office n'est **pas** une preuve qu'Excel peut être utilisé. La V4.1 conserve deux états distincts :

- **Licence Office** : état lu avec les outils Microsoft, selon l'édition.
- **Excel / test utilisateur** : `NON_VERIFIE`, `BLOQUE`, `FONCTIONNE`.

Le dernier test `FONCTIONNE` expire après 7 jours. Un signalement `BLOQUE` persiste jusqu'à confirmation manuelle de résolution. La surveillance n'ouvre pas Excel de force.

## Support WhatsApp

La V4.1 propose dans la **même fenêtre** un identifiant généré aléatoirement, par exemple `MON-0A1B2C3D4E5F`; il n'est pas dérivé du nom, du numéro de téléphone ou de la clé de licence.

Le bouton de support nécessite un consentement explicite. Il prépare **sur le Bureau** un ZIP contenant :

- `diagnostic.json` : version du moniteur, OS/Office, états, édition et **les 5 derniers caractères** seulement des clés ;
- `historique-expurge.txt` : historique de l'application, e-mails et clés complètes expurgés ;
- `evenements-office.json` : ID, date et fournisseur des erreurs applicatives ciblées, sans message complet ;
- `smart-app-control.json` : métadonnées des événements CodeIntegrity 3076 et 3077 ;
- `A-LIRE.txt` : rappel sur la confidentialité.

WhatsApp s'ouvre vers le numéro de support **+1 450 758-1093**, avec l'ID dans le texte prérempli. L'Explorateur sélectionne le ZIP. **L'utilisateur doit joindre le fichier puis appuyer lui-même sur Envoyer** : `wa.me` n'autorise pas le téléchargement/attachement automatique d'un document local via URL. Aucun envoi automatique de logs.

Le numéro WhatsApp de rappel est facultatif. Les données du rapport ne sont jamais envoyées en télémétrie.

## Statistiques locales

Des événements JSON minimalistes sont enregistrés localement pour les analyses, signalements de blocage Excel et rapports préparés. Ces événements ne contiennent ni nom, ni téléphone, ni clé et **ne sont pas envoyés vers un serveur**. Une future télémétrie centralisée exige une interface de consentement, un serveur de réception authentifié, une politique de rétention et une procédure de suppression.

## Blocage Smart App Control

Le fichier actuel est un lanceur **VBS/PowerShell non signé avec un certificat de code Authenticode**. La signature RSA du manifeste GitHub ne vaut pas signature Windows. Son exécution peut rester bloquée par Smart App Control. **Ne pas désactiver Smart App Control**.

Solution de distribution : application Windows native (par ex. .NET WPF empaquetée) et signature de code par une autorité approuvée, via Azure Artifact Signing / Microsoft Store MSIX, avec tests sur Windows 11 Smart App Control activé.

Références :
- https://learn.microsoft.com/windows/apps/develop/smart-app-control/code-signing-for-smart-app-control
- https://learn.microsoft.com/windows/apps/package-and-deploy/code-signing-options
- https://learn.microsoft.com/windows/apps/develop/smart-app-control/test-your-app-with-smart-app-control

## Recette Windows manuelle requise

1. Vérifier que le tableau de bord démarre et l'analyse se termine.
2. Avec Office actif et Excel non encore testé : **ne pas afficher un succès global vert**.
3. Tester le bouton **OUVRIR EXCEL** puis **EXCEL BLOQUÉ** : état rouge/alerte support après analyse.
4. Relancer le moniteur : le signalement reste présent.
5. Cliquer **EXCEL FONCTIONNE**, après avoir réellement modifié une feuille : statut confirmé, horodaté.
6. Sans consentement : **aucun rapport ZIP créé, aucun WhatsApp ouvert**.
7. Avec consentement : ZIP sur le Bureau, ID stable, conversation WhatsApp ouverte, **aucune pièce jointe envoyée automatiquement**.
8. Vérifier le contenu du ZIP : pas de clé complète, mot de passe ni chemin personnel dans les événements.
9. Vérifier que la surveillance quotidienne/hebdomadaire et les notifications de mise à jour gardent le comportement antérieur.
10. Tester le blocage et les événements CodeIntegrity 3077 sur une machine de test avec Smart App Control activé, puis procéder à la migration signée native.

La correction n'exécute pas Massgrave, ne force pas l'activation et ne contourne pas les contrôles Windows.
