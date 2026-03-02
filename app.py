"""
Morning Meeting Dashboard - Flask Server
Handles PPTX uploads, slide extraction, and serves the dashboard frontend.
"""

import os
import json
import shutil
import socket
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory, abort
from pptx_parser import extract_pptx_slides

# === Configuration ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
PUBLIC_DIR = os.path.join(BASE_DIR, 'public')
EXTRACTED_DIR = os.path.join(PUBLIC_DIR, 'extracted')
DATA_DIR = os.path.join(BASE_DIR, 'data')

PIC_DIR = os.path.join(DATA_DIR, 'pic_photos')

for d in [UPLOAD_DIR, EXTRACTED_DIR, DATA_DIR, PIC_DIR]:
    os.makedirs(d, exist_ok=True)

app = Flask(__name__, static_folder='public', static_url_path='')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

# === Default Departments ===
DEFAULT_DEPARTMENTS = {
    "departments": [
        {"id": "she", "name": "SHE", "icon": "🛡️"},
        {"id": "sales", "name": "SALES", "icon": "📊"},
        {"id": "delivery", "name": "DELIVERY", "icon": "🚚"},
        {"id": "subcont", "name": "SUBCONT", "icon": "🤝"},
        {"id": "qa", "name": "QA", "icon": "✅"},
        {"id": "plant1", "name": "PLANT 1", "icon": "🏭"},
        {"id": "plant2", "name": "PLANT 2", "icon": "🏗️"},
        {"id": "machine", "name": "MACHINE", "icon": "⚙️"},
        {"id": "dies_eng", "name": "DIES ENG", "icon": "🔧"},
    ]
}


def get_departments():
    config_path = os.path.join(DATA_DIR, 'departments.json')
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_DEPARTMENTS, f, indent=2, ensure_ascii=False)
        return DEFAULT_DEPARTMENTS


# ===================== ROUTES =====================

# --- Serve the SPA ---
@app.route('/')
def index():
    return send_from_directory(PUBLIC_DIR, 'index.html')


# --- API: Get departments ---
@app.route('/api/departments')
def api_departments():
    return jsonify(get_departments())


# --- API: Upload PPTX for a department ---
@app.route('/api/upload/<department>', methods=['POST'])
def api_upload(department):
    if 'pptx' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['pptx']
    if not file.filename.lower().endswith('.pptx'):
        return jsonify({"error": "Only .pptx files are allowed"}), 400

    # Save uploaded file
    dept_upload_dir = os.path.join(UPLOAD_DIR, department)
    os.makedirs(dept_upload_dir, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_name = f"{timestamp}_{file.filename}"
    file_path = os.path.join(dept_upload_dir, safe_name)
    file.save(file_path)

    # Clear previous extracted data
    dept_extract_dir = os.path.join(EXTRACTED_DIR, department)
    if os.path.exists(dept_extract_dir):
        shutil.rmtree(dept_extract_dir)
    os.makedirs(dept_extract_dir, exist_ok=True)

    try:
        result = extract_pptx_slides(file_path, dept_extract_dir, department)
        return jsonify({
            "success": True,
            "department": department,
            "filename": file.filename,
            "uploadedAt": datetime.now().isoformat(),
            "slides": result["slides"],
            "totalSlides": result["totalSlides"]
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- API: Get department slides ---
@app.route('/api/department/<department>/slides')
def api_dept_slides(department):
    meta_path = os.path.join(EXTRACTED_DIR, department, 'meta.json')
    if os.path.exists(meta_path):
        with open(meta_path, 'r', encoding='utf-8') as f:
            return jsonify(json.load(f))
    return jsonify({"slides": [], "totalSlides": 0, "message": "No presentation uploaded yet"})


# --- API: Delete department data ---
@app.route('/api/department/<department>', methods=['DELETE'])
def api_delete_dept(department):
    try:
        extract_dir = os.path.join(EXTRACTED_DIR, department)
        upload_dir = os.path.join(UPLOAD_DIR, department)
        if os.path.exists(extract_dir):
            shutil.rmtree(extract_dir)
        if os.path.exists(upload_dir):
            shutil.rmtree(upload_dir)
        return jsonify({"success": True, "message": f"Data for {department} deleted"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- API: Meeting info ---
@app.route('/api/meeting-info')
def api_meeting_info():
    now = datetime.now()
    return jsonify({
        "date": now.strftime('%A, %B %d, %Y'),
        "time": now.strftime('%I:%M %p')
    })


# --- Serve extracted files ---
@app.route('/extracted/<path:filename>')
def serve_extracted(filename):
    return send_from_directory(EXTRACTED_DIR, filename)


# --- Serve PIC photos ---
@app.route('/pic_photos/<path:filename>')
def serve_pic_photos(filename):
    return send_from_directory(PIC_DIR, filename)


# --- API: Get PIC data for a department ---
@app.route('/api/department/<department>/pic', methods=['GET'])
def api_get_pic(department):
    pic_path = os.path.join(DATA_DIR, f'pic_{department}.json')
    if os.path.exists(pic_path):
        with open(pic_path, 'r', encoding='utf-8') as f:
            return jsonify(json.load(f))
    return jsonify({'pics': []})


# --- API: Save PIC names (preserves existing photos) ---
@app.route('/api/department/<department>/pic', methods=['POST'])
def api_save_pic(department):
    data = request.get_json()
    incoming_pics = data.get('pics', [])

    pic_path = os.path.join(DATA_DIR, f'pic_{department}.json')
    existing = {'pics': []}
    if os.path.exists(pic_path):
        with open(pic_path, 'r', encoding='utf-8') as f:
            existing = json.load(f)

    existing_pics = existing.get('pics', [])
    merged = []
    for i, p in enumerate(incoming_pics):
        photo = existing_pics[i]['photo'] if i < len(existing_pics) else None
        merged.append({'name': p.get('name', ''), 'photo': photo})

    result = {'pics': merged}
    with open(pic_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return jsonify(result)


# --- API: Upload photo for a specific PIC slot ---
@app.route('/api/department/<department>/pic/photo/<int:index>', methods=['POST'])
def api_upload_pic_photo(department, index):
    if 'photo' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['photo']
    ext = os.path.splitext(file.filename)[1].lower() or '.jpg'
    if ext not in ['.jpg', '.jpeg', '.png', '.gif', '.webp']:
        return jsonify({'error': 'Invalid image format'}), 400

    # Remove any old photo for this slot
    for old_ext in ['.jpg', '.jpeg', '.png', '.gif', '.webp']:
        old_path = os.path.join(PIC_DIR, f'{department}_{index}{old_ext}')
        if os.path.exists(old_path):
            os.remove(old_path)

    filename = f'{department}_{index}{ext}'
    file.save(os.path.join(PIC_DIR, filename))
    photo_url = f'/pic_photos/{filename}'

    # Update the JSON file
    pic_path = os.path.join(DATA_DIR, f'pic_{department}.json')
    data = {'pics': []}
    if os.path.exists(pic_path):
        with open(pic_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

    pics = data.get('pics', [])
    while len(pics) <= index:
        pics.append({'name': '', 'photo': None})
    pics[index]['photo'] = photo_url
    data['pics'] = pics

    with open(pic_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return jsonify({'success': True, 'photo': photo_url})


# --- SPA fallback ---
@app.route('/<path:path>')
def spa_fallback(path):
    # Try to serve as static file first
    full_path = os.path.join(PUBLIC_DIR, path)
    if os.path.isfile(full_path):
        return send_from_directory(PUBLIC_DIR, path)
    return send_from_directory(PUBLIC_DIR, 'index.html')


# ===================== MAIN =====================
if __name__ == '__main__':
    # Detect LAN IP
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        lan_ip = s.getsockname()[0]
        s.close()
    except Exception:
        lan_ip = "localhost"

    print("\n========================================")
    print("  Morning Meeting Dashboard")
    print(f"  Local:   http://localhost:3000")
    print(f"  Network: http://{lan_ip}:3000")
    print("  Share the Network URL with other devices")
    print("========================================\n")
    app.run(host='0.0.0.0', port=3000, debug=True)
