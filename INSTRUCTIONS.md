# Propositions Commerciales Conwise Express

## Workflow pour générer une proposition client

### 1. Préparer les assets du client

Avant de lancer Claude, placer dans ce dossier :

- **Logo du client** en PNG (idéalement sans fond, utiliser remove.bg si besoin)
- **Screenshot du dashboard** personnalisé depuis conwise.app
- Noter l'**URL du site web** du client

### 2. Lancer Claude Code

```bash
cd "C:\Users\User\Documents\claude code\Conwise express propal"
claude
```

### 3. Donner le prompt

Adapter ce modèle selon le client :

```
Réalise une présentation PowerPoint pour [NOM DU CLIENT] pour la prise en charge
de leur flotte entreprise [+ préciser : agences France, services carrosserie, etc.].
Voici leur site : https://www.site-client.com/
C'est une entreprise de [secteur d'activité].
Utilise le cas client Derichebourg comme référence d'expertise multi-sites.
Le logo du client et le screenshot de leur dashboard sont dans le dossier.
```

### 4. Éléments personnalisés par client

| Élément | À adapter |
|---------|-----------|
| Logo client | Fichier PNG dans le dossier |
| Dashboard | Screenshot du portail conwise.app dédié |
| Enjeux flotte | Spécifiques au secteur et à l'organisation du client |
| Services proposés | Selon les besoins identifiés (VL, PL, carrosserie, etc.) |
| Nombre de sites/collaborateurs | Données issues du site web du client |

### 5. Ressources réutilisables (déjà dans le dossier)

- `Charte graphique.txt` — couleurs, polices, style Conwise Express
- `Argumentaire Gestionnaire de flotte.txt` — points clés de vente
- `Personas.txt` — profils types des décideurs ciblés
- Images génériques (convoyeur, état des lieux, remise en main, bannière)
- Logo Conwise Express

### 6. Scripts existants (exemples)

- `create_dmd_pptx.js` — Proposition Groupe DMD (distribution automobile)
- `generate_ortec.js` — Proposition Groupe ORTEC (ingénierie industrielle)

Ces scripts servent de base pour les nouvelles propositions.

## Clients réalisés

| Client | Secteur | Fichier |
|--------|---------|---------|
| Groupe DMD | Distribution automobile | `Proposition Conwise Express - Groupe DMD.pptx` |
| Groupe ORTEC | Ingénierie industrielle | `Proposition Conwise Express x Groupe ORTEC.pptx` |

## Contact Conwise Express

- Site : www.conwise-express.com
- Portail : conwise.app
- Email : contact@conwise-express.com
- Tél : +33 7 44 31 79 16
