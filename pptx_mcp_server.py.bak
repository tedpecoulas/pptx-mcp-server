"""
Claude Skills MCP Server - PPTX Edition
Python server for reading and modifying PowerPoint templates

Required libraries:
pip install python-pptx requests flask flask-cors

Deploy on: Railway.app, Render.com, or Fly.io
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import requests
import io
import json
import tempfile
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)

# Store modified presentations temporarily
temp_files = {}


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


@app.route('/api/mcp', methods=['GET', 'POST', 'OPTIONS'])
def mcp_endpoint():
    """Main MCP endpoint following JSON-RPC 2.0"""
    
    if request.method == 'OPTIONS':
        return '', 200
    
    if request.method == 'GET':
        return jsonify({
            "name": "PPTX MCP Server",
            "version": "1.0.0",
            "tools": ["analyze_template", "modify_template", "list_tools"]
        })
    
    # Handle POST - JSON-RPC
    body = request.get_json()
    request_id = body.get('id', 1)
    method = body.get('method', '')
    params = body.get('params', {})
    
    print(f"📥 Method: {method}")
    print(f"📥 Params: {json.dumps(params, indent=2)}")
    
    # Route: initialize
    if method == 'initialize':
        client_version = params.get('protocolVersion', '2025-06-18')
        return jsonify({
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
        })
    
    # Route: tools/list
    if method == 'tools/list':
        return jsonify({
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
                                }
                            },
                            "required": ["template_url", "modifications"]
                        }
                    }
                ]
            }
        })
    
    # Route: tools/call
    if method == 'tools/call':
        tool_name = params.get('name')
        args = params.get('arguments', {})
        
        # Tool: analyze_template
        if tool_name == 'analyze_template':
            try:
                template_url = args.get('template_url')
                
                print(f"📄 Analyzing template: {template_url}")
                
                # Download and analyze
                prs = download_pptx(template_url)
                analysis = analyze_presentation(prs)
                
                return jsonify({
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "content": [{
                            "type": "text",
                            "text": json.dumps(analysis, indent=2, ensure_ascii=False)
                        }]
                    }
                })
            
            except Exception as e:
                return jsonify({
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32603,
                        "message": f"Error analyzing template: {str(e)}"
                    }
                })
        
        # Tool: modify_template
        if tool_name == 'modify_template':
            try:
                template_url = args.get('template_url')
                modifications = args.get('modifications', {})
                
                print(f"✏️ Modifying template: {template_url}")
                print(f"✏️ Modifications: {json.dumps(modifications, indent=2)}")
                
                # Download template
                prs = download_pptx(template_url)
                
                # Apply modifications
                prs = modify_presentation(prs, modifications)
                
                # Save to temporary file
                output_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
                prs.save(output_file.name)
                
                # Store file path with timestamp
                file_id = f"pptx_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                temp_files[file_id] = output_file.name
                
                # Generate download URL
                download_url = f"/download/{file_id}"
                
                return jsonify({
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "content": [{
                            "type": "text",
                            "text": f"✅ Template modifié avec succès!\n\nTélécharger: {download_url}\n\nModifications appliquées: {len(modifications)} slides"
                        }]
                    }
                })
            
            except Exception as e:
                return jsonify({
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32603,
                        "message": f"Error modifying template: {str(e)}"
                    }
                })
        
        # Unknown tool
        return jsonify({
            "jsonrpc": "2.0",
            "id": request_id,
            "error": {
                "code": -32601,
                "message": f"Unknown tool: {tool_name}"
            }
        })
    
    # Unknown method
    return jsonify({
        "jsonrpc": "2.0",
        "id": request_id,
        "error": {
            "code": -32601,
            "message": f"Method not found: {method}"
        }
    })


@app.route('/download/<file_id>')
def download_file(file_id):
    """Download endpoint for modified presentations"""
    if file_id not in temp_files:
        return jsonify({"error": "File not found"}), 404
    
    file_path = temp_files[file_id]
    
    if not os.path.exists(file_path):
        return jsonify({"error": "File no longer exists"}), 404
    
    return send_file(
        file_path,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        as_attachment=True,
        download_name=f'modified_{file_id}.pptx'
    )


@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({
        "status": "healthy",
        "server": "pptx-mcp-server",
        "version": "1.0.0"
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
