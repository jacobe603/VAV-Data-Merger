"""
VAV Data Merger - Main Application

A web-based tool for combining Excel spreadsheet data with Titus Teams TW2 database files.
Refactored into modular architecture for better maintainability.
"""
import os
import logging
from flask import Flask
from flask_cors import CORS
from flask_session import Session

# Import configuration
from models.config import (
    SECRET_KEY, DEBUG, PORT,
    SESSION_TYPE, SESSION_FILE_DIR, SESSION_PERMANENT,
    SESSION_USE_SIGNER, SESSION_KEY_PREFIX,
    UPLOAD_FOLDER, MAX_CONTENT_LENGTH
)

# Import route blueprints
from routes.main_routes import main_bp
from routes.tw2_routes import tw2_bp
from routes.excel_routes import excel_bp
from routes.comparison_routes import comparison_bp
from routes.debug_routes import debug_bp


def create_app():
    """Application factory function."""
    app = Flask(__name__)

    # Configure app
    app.secret_key = SECRET_KEY
    app.config['SESSION_TYPE'] = SESSION_TYPE
    app.config['SESSION_FILE_DIR'] = SESSION_FILE_DIR
    app.config['SESSION_PERMANENT'] = SESSION_PERMANENT
    app.config['SESSION_USE_SIGNER'] = SESSION_USE_SIGNER
    app.config['SESSION_KEY_PREFIX'] = SESSION_KEY_PREFIX
    app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
    app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

    # Initialize extensions
    Session(app)
    CORS(app)

    # Setup logging
    logger = logging.getLogger('vav_data_merger')
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
        logger.addHandler(handler)

    # Ensure required directories exist
    os.makedirs(SESSION_FILE_DIR, exist_ok=True)
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

    # Register blueprints
    app.register_blueprint(main_bp)
    app.register_blueprint(tw2_bp)
    app.register_blueprint(excel_bp)
    app.register_blueprint(comparison_bp)
    app.register_blueprint(debug_bp)

    return app


if __name__ == '__main__':
    app = create_app()
    print("=" * 70)
    print("VAV Data Merger - Starting Application")
    print("=" * 70)
    print(f"Running on: http://127.0.0.1:{PORT}")
    print(f"Debug mode: {DEBUG}")
    print("=" * 70)
    app.run(debug=DEBUG, port=PORT)
