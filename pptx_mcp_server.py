"""
Claude Skills MCP Server - PPTX Edition with SSE Support
Python server for reading and modifying PowerPoint templates
Supports both JSON-RPC and SSE transports for DUST compatibility
"""

from flask import Flask, request, jsonify, send_file, Response
from flask_cors import CORS
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
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


def sanitize_filename(text):
    """Sanitize text for use in filename"""
    # Remove or replace invalid characters
    text = re.sub(r'[<>:"/\\|?*]', '-', text)
    # Remove leading/trailing spaces and dots
    text = text.strip(' .')
    # Limit length
    return text[:50] if text else "Document"


def download_pptx(url):
    """Download PPTX from URL and return Presentation object"""
    response = requests.get(url)
    response.raise_for_status()
    pptx_bytes = io.BytesIO(response.content)
    return Presentation(pptx_bytes)


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
                "has_text_frame": shape.has_text_frame
            }
            
            # Extract text if available
            if shape.has_text_frame:
                text = shape.text_frame.text
                shape_info["text"] = text
                shape_info["text_length"] = len(text)
                
                # Check if it's a placeholder
                if shape.is_placeholder:
                    shape_info["placeholder_type"] = str(shape.placeholder_format.type)
                else:
                    shape_info["placeholder_type"] = None
                
                # Count paragraphs
                shape_info["paragraph_count"] = len(shape.text_frame.paragraphs)
            
            # Check if it's a picture
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shape_info["is_picture"] = True
            
            slide_info["shapes"].append(shape_info)
        
        analysis["slides"].append(slide_info)
    
    return analysis


def modify_presentation(prs, modifications):
    """Modify presentation based on modifications dict"""
    for slide_key, shape_mods in modifications.items():
        # Extract slide number from "slide_0" format
        slide_num = int(slide_key.split('_')[1])
        
        if slide_num >= len(prs.slides):
            continue
        
        slide = prs.slides[slide_num]
        
        for shape_key, new_text in shape_mods.items():
            # Extract shape number from "shape_0" format
            shape_num = int(shape_key.split('_')[1])
            
            if shape_num >= len(slide.shapes):
                continue
            
            shape = slide.shapes[shape_num]
            
            if shape.has_text_frame:
                # Replace text while preserving formatting
                shape.text_frame.text = new_text
    
    return prs


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
                    "version": "1.0.0"
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
                        "description": "Modifie un template PowerPoint en remplaçant le texte dans les zones identifiées",
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
                                    "description": "Métadonnées optionnelles pour nommer le fichier (client, mission, consultant)",
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
                prs = modify_presentation(prs, modifications)
                
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
                
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "content": [{
                            "type": "text",
                            "text": f"✅ Votre REX est prêt !\n\n📥 Télécharger ici: {download_url}\n\n💡 Nom de fichier: {suggested_name}\n\n(Le fichier se téléchargera avec ce nom automatiquement)"
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
            "version": "1.0.0",
            "tools": ["analyze_template", "modify_template"]
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
            # Process the request
            response_data = handle_mcp_request(body, request_id)
            
            # Send as SSE event
            sse_data = f"data: {json.dumps(response_data)}\n\n"
            yield sse_data
            
            # Keep connection alive briefly
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
        "version": "1.0.0",
        "transport": "JSON + SSE"
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
