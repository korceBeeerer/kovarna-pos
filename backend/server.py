from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from excel_handler import (add_order_to_excel, storno_order_in_excel, load_products_from_excel,
                           backup_excel, get_backup_list, get_config, set_excel_file)
from datetime import datetime
import os

app = Flask(__name__)
CORS(app)  # Povolí přístup z tabletu

# Cesta k frontend složce
FRONTEND_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'frontend')

@app.route('/')
def serve_frontend():
    """Servíruje hlavní stránku aplikace"""
    return send_from_directory(FRONTEND_DIR, 'index.html')

@app.route('/<path:path>')
def serve_static(path):
    """Servíruje statické soubory (CSS, JS, obrázky)"""
    return send_from_directory(FRONTEND_DIR, path)

@app.route('/api/products', methods=['GET'])
def get_products():
    """
    Vrátí seznam produktů z Excel souboru
    """
    try:
        products = load_products_from_excel()
        return jsonify({
            "success": True,
            "products": products
        }), 200
    except Exception as e:
        print(f"Chyba při načítání produktů: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

@app.route('/api/order', methods=['POST'])
def save_order():
    """
    Přijme objednávku z tabletu a uloží ji do Excelu

    Očekávaná data:
    {
        "items": [
            {"name": "Produkt 1", "price": 100, "quantity": 2},
            {"name": "Produkt 2", "price": 50, "quantity": 1}
        ],
        "total": 250
    }
    """
    try:
        data = request.json

        if not data or 'items' not in data or 'total' not in data:
            return jsonify({"error": "Neplatná data"}), 400

        # Přidej časovou značku
        order_data = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "items": data['items'],
            "total": data['total']
        }

        # Uložit do Excelu
        result = add_order_to_excel(order_data)

        if result['success']:
            return jsonify(result), 200
        else:
            return jsonify(result), 500

    except Exception as e:
        print(f"Chyba: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

@app.route('/api/storno', methods=['POST'])
def storno_order():
    """
    Stornuje položky - odečte je z Excelu pro aktuální den.
    Očekávaná data jsou stejná jako u /api/order.
    """
    try:
        data = request.json

        if not data or 'items' not in data or 'total' not in data:
            return jsonify({"error": "Neplatná data"}), 400

        order_data = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "items": data['items'],
            "total": data['total']
        }

        result = storno_order_in_excel(order_data)

        if result['success']:
            return jsonify(result), 200
        else:
            return jsonify(result), 400

    except Exception as e:
        print(f"Chyba: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

@app.route('/api/config', methods=['GET'])
def read_config():
    """Vrátí aktuální nastavení (cestu k Excel souboru)"""
    return jsonify({"success": True, **get_config()}), 200

@app.route('/api/config', methods=['POST'])
def update_config():
    """
    Nastaví novou cestu k Excel souboru.

    Očekávaná data:
    {"excel_file": "C:\\Kovarna\\Tržby.xlsx"}
    """
    data = request.json or {}
    result = set_excel_file(data.get('excel_file'))
    if result['success']:
        return jsonify({**result, **get_config()}), 200
    else:
        return jsonify(result), 400

@app.route('/api/health', methods=['GET'])
def health_check():
    """Kontrola, zda server běží"""
    return jsonify({"status": "ok", "message": "Server běží"}), 200

@app.route('/api/backup', methods=['POST'])
def create_backup():
    """
    Vytvoří manuální zálohu Excel souboru
    """
    try:
        result = backup_excel()
        if result['success']:
            return jsonify(result), 200
        else:
            return jsonify(result), 500
    except Exception as e:
        print(f"Chyba při vytváření zálohy: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

@app.route('/api/backups', methods=['GET'])
def list_backups():
    """
    Vrátí seznam existujících záloh
    """
    try:
        backups = get_backup_list()
        return jsonify({
            "success": True,
            "backups": backups,
            "count": len(backups)
        }), 200
    except Exception as e:
        print(f"Chyba při načítání záloh: {str(e)}")
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

if __name__ == '__main__':
    # Server poběží na všech síťových rozhraních (přístupný z tabletu)
    # Port 5000
    print("=" * 50)
    print("KOVÁRENSKÝ SYSTÉM - SERVER")
    print("=" * 50)
    print("\nServer běží na:")
    print("  - Lokálně: http://localhost:5000")
    print("  - V síti: http://<IP_adresa_PC>:5000")
    print("\nPro zastavení serveru stiskněte Ctrl+C")
    print("=" * 50)

    app.run(host='0.0.0.0', port=5000, debug=True)
