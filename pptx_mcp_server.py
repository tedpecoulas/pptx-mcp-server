"""
Claude Skills MCP Server - PPTX Edition with SSE Support
Python server for reading and modifying PowerPoint templates
Supports both JSON-RPC and SSE transports for DUST compatibility
WITH INTELLIGENT FONT AUTO-SIZING
"""

from flask import Flask, request, jsonify, send_file, Response
from flask_cors import CORS
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
from pptx.enum.text import MSO_AUTO_SIZE
import requests
import io
import json
import tempfile
import os
import time
from datetime import datetime
import re

app = Flask(__name__)
CORS(app)

# Store modified presentations temporarily
temp_files = {}

# Configuration des cadres à formater uniformément
UNIFORM_FORMAT_SHAPES = [
    "contexte",
    "travaux réalisés", 
    "type de mission",
    "outils utilisés",
    "résultats"
]

# Taille de police par défaut et minimale
DEFAULT_FONT_SIZE = 12
MIN_FONT_SIZE = 8


def sanitize_filename(text):
    """Sanitize text for use in filename"""
    text = re.sub(r'[<>:"/\\|?*]', '-', text)
    text = text.strip(' .')
    return text[:50] if text else "Document"


def download_pptx(url):
    """Download PPTX from URL and return Presentation object"""
    response = requests.get(url)
    response.raise_for_status()
    pptx_bytes = io.BytesIO(response.content)
    return Presentation(pptx_bytes)


def normalize_shape_name(name):
    """Normalise le nom d'une shape pour comparaison"""
    return name.lower().strip()


def is_uniform_format_shape(shape):
    """Vérifie si une shape fait partie des cadres à formater uniformément"""
    if not shape.has_text_frame:
        return False
    
    shape_name_normalized = normalize_shape_name(shape.name)
    
    # Vérifier si le nom de la shape contient un des mots-clés
    for keyword in UNIFORM_FORMAT_SHAPES:
        if keyword.lower() in shape_name_normalized:
            return True
    
    # Vérifier aussi le texte actuel de la shape (cas des placeholders)
    if shape.text_frame.text:
        text_normalized = normalize_shape_name(shape.text_frame.text)
        for keyword in UNIFORM_FORMAT_SHAPES:
            if keyword.lower() in text_normalized:
                return True
    
    return False


def calculate_optimal_font_size(texts, max_size=DEFAULT_FONT_SIZE, min_size=MIN_FONT_SIZE):
    """
    Calcule la taille de police optimale pour plusieurs textes
    en se basant sur le texte le plus long
    """
    if not texts:
        return max_size
    
    # Trouver la longueur maximale
    max_length = max(len(text) for text in texts if text)
    
    # Calculer la taille optimale selon la longueur
    if max_length < 100:
        font_size = max_size
    elif max_length < 200:
        font_size = max_size - 1
    elif max_length < 300:
        font_size = max_size - 2
    elif max_length < 500:
        font_size = max_size - 3
    else:
        font_size = min_size
    
    # S'assurer de ne pas descendre sous la taille minimale
    font_size = max(font_size, min_size)
    
    print(f"📏 [FONT-CALC] Max length: {max_length} chars → Font size: {font_size}pt")
    return font_size


def apply_font_size_to_shape(shape, text, font_size):
    """Applique une taille de police à une shape"""
    if not shape.has_text_frame:
        return False
    
    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.auto_size = MSO_AUTO_SIZE.NONE
    
    # Ajouter le texte
    p = text_frame.paragraphs[0]
    p.text = text
    
    # Appliquer la taille de police à tous les runs
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
    
    print(f"✍️  [APPLY-FONT] Shape '{shape.name}': {len(text)} chars → {font_size}pt")
    return True


def analyze_presentation(prs):
    """Analyze presentation structure and return JSON"""
    analysis = {
        "total_slides": len(prs.slides),
        "slides": []
    }
    
    for slide_idx, slide in enumerate(prs.slides):
        slide_info = {
            "slide_number": slide_idx,
            "layout_name": slide.slide_layout.name,
            "shapes": []
        }
        
        for shape_idx, shape in enumerate(slide.shapes):
            shape_info = {
                "shape_id": shape_idx,
                "name": shape.name,
                "type": str(shape.shape_type),
                "has_text_frame": shape.has_text_frame,
                "is_uniform_format": is_uniform_format_shape(shape)
            }
            
            if shape.has_text_frame:
                text = shape.text_frame.text
                shape_info["text"] = text
                shape_info["text_length"] = len(text)
                
                if shape.is_placeholder:
                    shape_info["placeholder_type"] = str(shape.placeholder_format.type)
                else:
                    shape_info["placeholder_type"] = None
                
                shape_info["paragraph_count"] = len(shape.text_frame.paragraphs)
            
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shape_info["is_picture"] = True
            
            slide_info["shapes"].append(shape_info)
        
        analysis["slides"].append(slide_info)
    
    return analysis


def modify_presentation(prs, modifications):
    """
    Modifie la présentation avec ajustement intelligent de la police
    """
    warnings = []
    
    # Phase 1 : Identifier les shapes uniformes et leurs textes
    uniform_shapes_data = []
    
    for slide_key, shape_mods in modifications.items():
        slide_num = int(slide_key.split('_')[1])
        
        if slide_num >= len(prs.slides):
            continue
        
        slide = prs.slides[slide_num]
        
        for shape_key, new_text in shape_mods.items():
            shape_num = int(shape_key.split('_')[1])
            
            if shape_num >= len(slide.shapes):
                continue
            
            shape = slide.shapes[shape_num]
            
            if is_uniform_format_shape(shape):
                uniform_shapes_data.append({
                    'shape': shape,
                    'text': new_text,
                    'slide_num': slide_num,
                    'shape_num': shape_num
                })
    
    # Phase 2 : Calculer la taille de police optimale pour les shapes uniformes
    uniform_font_size = DEFAULT_FONT_SIZE
    if uniform_shapes_data:
        uniform_texts = [data['text'] for data in uniform_shapes_data]
        uniform_font_size = calculate_optimal_font_size(uniform_texts)
        
        print(f"🎯 [UNIFORM] {len(uniform_shapes_data)} shapes with uniform font: {uniform_font_size}pt")
        
        # Vérifier si on est à la taille minimale
        if uniform_font_size == MIN_FONT_SIZE:
            max_length = max(len(text) for text in uniform_texts)
            warnings.append(
                f"⚠️ ATTENTION : Un ou plusieurs cadres (Contexte, Travaux, etc.) "
                f"contiennent beaucoup de texte ({max_length} caractères max). "
                f"La police a été réduite au minimum ({MIN_FONT_SIZE}pt). "
                f"Pour une meilleure lisibilité, réduisez le contenu à ~300-400 caractères."
            )
    
    # Phase 3 : Appliquer les modifications
    for slide_key, shape_mods in modifications.items():
        slide_num = int(slide_key.split('_')[1])
        
        if slide_num >= len(prs.slides):
            continue
        
        slide = prs.slides[slide_num]
        
        for shape_key, new_text in shape_mods.items():
            shape_num = int(shape_key.split('_')[1])
            
            if shape_num >= len(slide.shapes):
                continue
            
            shape = slide.shapes[shape_num]
            
            if not shape.has_text_frame:
                continue
            
            # Appliquer la taille de police appropriée
            if is_uniform_format_shape(shape):
                # Shapes uniformes : même taille pour toutes
                apply_font_size_to_shape(shape, new_text, uniform_font_size)
            else:
                # Autres shapes : calcul individuel
                individual_font_size = calculate_optimal_font_size([new_text])
                apply_font_size_to_shape(shape, new_text, individual_font_size)
    
    return prs, warnings


def handle_mcp_request(body, request_id):
    """Handle MCP JSON-RPC request and return response"""
    method = body.get('method', '')
    params = body.get('params', {})
    
    print(f"📥 Method: {method}")
    
    # Route: initialize
    if method == 'initialize':
        client_version = params.get('protocolVersion', '2025-06-18')
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "protocolVersion": client_version,
                "capabilities": {
                    "tools": {"listChanged": False},
                    "resources": {},
                    "prompts": {}
                },
                "serverInfo": {
                    "name": "pptx-mcp-server",
                    "version": "2.0.0"
                }
            }
        }
    
    # Route: tools/list
    if method == 'tools/list':
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "tools": [
                    {
                        "name": "analyze_template",
                        "description": "Analyse la structure d'un template PowerPoint (slides, zones de texte, images)",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "template_url": {
                                    "type": "string",
                                    "description": "URL du fichier PPTX à analyser"
                                }
                            },
                            "required": ["template_url"]
                        }
                    },
                    {
                        "name": "modify_template",
                        "description": "Modifie un template PowerPoint avec ajustement automatique de la police pour éviter les débordements",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "template_url": {
                                    "type": "string",
                                    "description": "URL du template PPTX"
                                },
                                "modifications": {
                                    "type": "object",
                                    "description": "Dictionnaire des modifications (slide_X: {shape_Y: nouveau_texte})"
                                },
                                "metadata": {
                                    "type": "object",
                                    "description": "Métadonnées pour nommer le fichier (client, mission, consultant)",
                                    "properties": {
                                        "client": {"type": "string"},
                                        "mission": {"type": "string"},
                                        "consultant": {"type": "string"}
                                    }
                                }
                            },
                            "required": ["template_url", "modifications"]
                        }
                    }
                ]
            }
        }
    
    # Route: tools/call
    if method == 'tools/call':
        tool_name = params.get('name')
        args = params.get('arguments', {})
        
        # Tool: analyze_template
        if tool_name == 'analyze_template':
            try:
                template_url = args.get('template_url')
                print(f"📄 Analyzing template: {template_url}")
                
                prs = download_pptx(template_url)
                analysis = analyze_presentation(prs)
                
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "content": [{
                            "type": "text",
                            "text": json.dumps(analysis, indent=2, ensure_ascii=False)
                        }]
                    }
                }
            except Exception as e:
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32603,
                        "message": f"Error analyzing template: {str(e)}"
                    }
                }
        
        # Tool: modify_template
        if tool_name == 'modify_template':
            try:
                template_url = args.get('template_url')
                modifications = args.get('modifications', {})
                metadata = args.get('metadata', {})
                
                print(f"✏️ Modifying template: {template_url}")
                print(f"✏️ Metadata: {metadata}")
                
                prs = download_pptx(template_url)
                prs, warnings = modify_presentation(prs, modifications)
                
                # Generate filename from metadata
                client = sanitize_filename(metadata.get('client', ''))
                mission = sanitize_filename(metadata.get('mission', ''))
                consultant = sanitize_filename(metadata.get('consultant', ''))
                
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_id = f"pptx_{timestamp}"
                
                # Build suggested filename
                if client and mission and consultant:
                    suggested_name = f"REX - {client} - {mission} - {consultant}.pptx"
                elif client and mission:
                    suggested_name = f"REX - {client} - {mission}.pptx"
                elif client:
                    suggested_name = f"REX - {client}.pptx"
                else:
                    suggested_name = f"REX_{timestamp}.pptx"
                
                # Save file
                output_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
                prs.save(output_file.name)
                
                temp_files[file_id] = {
                    'path': output_file.name,
                    'suggested_name': suggested_name
                }
                
                # Construct full URL
                base_url = os.environ.get('SERVER_URL', 'https://pptx-mcp-server-production.up.railway.app')
                download_url = f"{base_url}/download/{file_id}"
                
                # Construire le message de réponse
                response_text = f"✅ Votre REX est prêt !\n\n📥 Télécharger ici: {download_url}\n\n💡 Nom de fichier: {suggested_name}\n\n"
                
                # Ajouter les warnings si présents
                if warnings:
                    response_text += "\n" + "\n\n".join(warnings)
                
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "content": [{
                            "type": "text",
                            "text": response_text
                        }]
                    }
                }
            except Exception as e:
                print(f"❌ Error: {str(e)}")
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32603,
                        "message": f"Error modifying template: {str(e)}"
                    }
                }
        
        # Unknown tool
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "error": {
                "code": -32601,
                "message": f"Unknown tool: {tool_name}"
            }
        }
    
    # Unknown method
    return {
        "jsonrpc": "2.0",
        "id": request_id,
        "error": {
            "code": -32601,
            "message": f"Method not found: {method}"
        }
    }


@app.route('/api/mcp', methods=['GET', 'POST', 'OPTIONS'])
def mcp_endpoint():
    """Main MCP endpoint - supports both JSON and SSE transports"""
    
    if request.method == 'OPTIONS':
        return '', 200
    
    if request.method == 'GET':
        return jsonify({
            "name": "PPTX MCP Server",
            "version": "2.0.0",
            "tools": ["analyze_template", "modify_template"],
            "features": ["intelligent_font_sizing", "uniform_format_shapes"]
        })
    
    # Handle POST - check if client wants SSE
    accept_header = request.headers.get('Accept', '')
    wants_sse = 'text/event-stream' in accept_header
    
    print(f"📥 Accept header: {accept_header}")
    print(f"📥 Wants SSE: {wants_sse}")
    
    body = request.get_json() or {}
    request_id = body.get('id', 1)
    
    # If SSE is requested, use SSE transport
    if wants_sse:
        print("🔄 Using SSE transport")
        
        def generate_sse():
            response_data = handle_mcp_request(body, request_id)
            sse_data = f"data: {json.dumps(response_data)}\n\n"
            yield sse_data
            time.sleep(0.5)
        
        return Response(
            generate_sse(),
            mimetype='text/event-stream',
            headers={
                'Cache-Control': 'no-cache',
                'X-Accel-Buffering': 'no',
                'Connection': 'keep-alive'
            }
        )
    
    # Otherwise use standard JSON response
    print("📤 Using JSON transport")
    response_data = handle_mcp_request(body, request_id)
    return jsonify(response_data)


@app.route('/download/<file_id>')
def download_file(file_id):
    """Download endpoint for modified presentations"""
    if file_id not in temp_files:
        return jsonify({"error": "File not found"}), 404
    
    file_info = temp_files[file_id]
    file_path = file_info['path']
    suggested_name = file_info['suggested_name']
    
    if not os.path.exists(file_path):
        return jsonify({"error": "File no longer exists"}), 404
    
    return send_file(
        file_path,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        as_attachment=True,
        download_name=suggested_name
    )


@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({
        "status": "healthy",
        "server": "pptx-mcp-server",
        "version": "2.0.0",
        "transport": "JSON + SSE",
        "features": {
            "intelligent_font_sizing": True,
            "uniform_format_shapes": UNIFORM_FORMAT_SHAPES,
            "min_font_size": MIN_FONT_SIZE,
            "default_font_size": DEFAULT_FONT_SIZE
        }
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)