# Moniteur Office + Windows

**Projet de développement (V4.0.0, non encore validé sur un ordinateur Windows réel).**

Tableau de bord Windows / Microsoft Office, diagnostic des licences, suivi facultatif quotidien ou hebdomadaire, interventions officielles Microsoft uniquement après confirmation.

## Installation

Un seul fichier `Moniteur-Office-Windows.vbs` à placer sur le Bureau et à ouvrir par double-clic. L'application copie ses composants dans `%LOCALAPPDATA%\ControleLicencesOfficeWindows` afin d'assurer la continuité de la tâche planifiée.

**Important :** les scripts Windows Script Host peuvent être bloqués par les politiques de sécurité d'une organisation. Ne désactivez pas Defender ou SmartScreen pour contourner un blocage.

## Mises à jour

- Vérification du manifeste `updates/latest.json` à l'ouverture du tableau de bord et lors des vérifications planifiées actives.
- Signature RSA-SHA256 du manifeste, avec clé publique intégrée au moniteur et clé privée conservée hors GitHub.
- Vérification SHA-256 des octets téléchargés **avant** remplacement du programme local.
- Nouvelle version proposée dans la fenêtre ; **aucune installation sans réponse positive de l'utilisateur**.
- Bouton « Plus tard » : reporter le rappel de sept jours.
- Sauvegarde locale de la version précédente et restauration si la copie ou l'intégrité échoue.
- Le fichier du Bureau de la V4 ne peut pas écraser une version plus récente déjà installée.

Le manifeste signé contient : `version`, `url`, `sha256`, `notes`, `signature`. La signature couvre exactement le texte UTF-8 `version + LF + url + LF + sha256 (minuscules) + LF + notes + LF`.

**Publication :** téléverser la nouvelle version, calculer SHA-256, signer le manifeste avec la clé privée hors ligne, puis mettre à jour `updates/latest.json` **en dernier**. Publier et tester l'installation sur une machine virtuelle avant de la distribuer.

## Limitations

- La V3 n'a pas de moteur de mise à jour : tous les postes existants doivent recevoir **une fois** la V4.
- Les vérifications planifiées sans session Windows ouverte n'affichent pas d'invite ; elles peuvent être remises au prochain lancement selon la configuration Windows.
- Les dates de dernière activation et les clés complètes ne sont pas accessibles de façon fiable dans les outils Microsoft. Le moniteur ne fabrique pas ces données.
- Ce projet n'intègre pas les contournements de licence Massgrave. Les tentatives de réactivation utilisent seulement les outils Microsoft et les licences déjà présentes.
- **La V4 est une candidate, non validée de bout en bout sur Windows ; ne pas l'envoyer à un parc d'utilisateurs avant validation.**

## Sécurité des publications

Ne jamais ajouter la clé privée de signature, des clés de licences, des données de clients ou des jetons GitHub dans ce dépôt public.
