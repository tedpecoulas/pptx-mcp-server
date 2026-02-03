# PPTX MCP Server

Serveur MCP pour analyser et modifier des templates PowerPoint.

## Outils Disponibles

### 1. `analyze_template`
Analyse la structure complète d'un template PPTX :
- Nombre de slides
- Zones de texte par slide
- Contenu actuel de chaque zone
- Type de placeholder

**Input:**
```json
{
  "template_url": "https://example.com/template.pptx"
}
```

### 2. `modify_template`
Modifie un template PPTX en conservant le format :
- Remplace le texte dans les zones identifiées
- Conserve tout le formatage
- Génère un nouveau fichier téléchargeable

**Input:**
```json
{
  "template_url": "https://example.com/template.pptx",
  "modifications": {
    "slide_0": {
      "shape_0": "Nouveau titre",
      "shape_1": "Nouveau sous-titre"
    },
    "slide_1": {
      "shape_1": "• Point 1\n• Point 2\n• Point 3"
    }
  }
}
```

## Déploiement

1. Push sur GitHub
2. Connectez Railway.app à votre repo
3. Railway déploie automatiquement
4. Utilisez l'URL : `https://votre-url.railway.app/api/mcp`

## Endpoints

- `GET /health` - Health check
- `GET /api/mcp` - Info serveur
- `POST /api/mcp` - Endpoint MCP principal
- `GET /download/{file_id}` - Télécharger fichier modifié

## Configuration DUST

**Dans DUST Admin:**
- Name: `PPTX Editor`
- URL: `https://votre-url.railway.app/api/mcp`
- Authentication: None

Le serveur exposera 2 outils que vos agents DUST pourront utiliser.

## Développement Local

```bash
pip install -r requirements.txt
python pptx_mcp_server.py
```

Testez : `http://localhost:5000/health`
