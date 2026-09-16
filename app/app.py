
import os
import html
import imaplib
import smtplib
import email
from email.message import EmailMessage
from email.header import decode_header
import sqlite3
import re
import socket
import ipaddress
import io
import csv
import json
import shutil
import zipfile
import urllib.request
import urllib.parse
import urllib.error
import statistics
from datetime import datetime, date, timedelta
from functools import wraps
from pathlib import Path

from flask import Flask, Response, flash, jsonify, redirect, render_template, render_template_string, request, send_file, send_from_directory, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix
from openpyxl import load_workbook

try:
    import qrcode
    import qrcode.image.svg
except ImportError:
    qrcode = None

APP_DIR = Path(__file__).resolve().parent
PROJECT_DIR = APP_DIR.parent
BASE_DIR = PROJECT_DIR

# Azure App Service keeps /home persistent between restarts. Local Windows use
# continues to default to the project folders, so the same build works both ways.
DATA_ROOT = Path(os.environ.get("BAM_DATA_ROOT", str(PROJECT_DIR))).expanduser().resolve()
DB_DIR = Path(os.environ.get("BAM_DB_DIR", str(DATA_ROOT / "database"))).expanduser().resolve()
UPLOAD_DIR = Path(os.environ.get("BAM_UPLOAD_DIR", str(DATA_ROOT / "uploads"))).expanduser().resolve()
BACKUP_DIR = Path(os.environ.get("BAM_BACKUP_DIR", str(DATA_ROOT / "backups"))).expanduser().resolve()
REPORT_DIR = Path(os.environ.get("BAM_REPORT_DIR", str(DATA_ROOT / "reports"))).expanduser().resolve()

for folder in (DB_DIR, UPLOAD_DIR, BACKUP_DIR, REPORT_DIR):
    folder.mkdir(parents=True, exist_ok=True)

DB_PATH = Path(os.environ.get("BAM_SQLITE_PATH", str(DB_DIR / "bam_motor_group.db"))).expanduser().resolve()

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1)
app.secret_key = os.environ.get("BAM_SECRET_KEY", "change-this-secret-key-before-production")
app.config.update(
    MAX_CONTENT_LENGTH=int(os.environ.get("BAM_MAX_UPLOAD_MB", "100")) * 1024 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("BAM_SECURE_COOKIES", "0") == "1",
    PERMANENT_SESSION_LIFETIME=timedelta(hours=int(os.environ.get("BAM_SESSION_HOURS", "12"))),
)

APP_VERSION = "25.18.7"
APP_NAME = "BAM Dealer Enterprise Cloud"
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-5").strip() or "gpt-5"
VIN_DECODER_URL = os.environ.get("BAM_VIN_DECODER_URL", "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValuesExtended/{vin}?format=json")
VIN_DATA_PROVIDER_NAME = os.environ.get("BAM_VIN_DATA_PROVIDER_NAME", "NHTSA vPIC").strip() or "NHTSA vPIC"
WORKSHOP_PROVIDER_NAME = os.environ.get("BAM_WORKSHOP_PROVIDER_NAME", "eManualOnline").strip() or "eManualOnline"
WORKSHOP_PORTAL_URL = os.environ.get("BAM_WORKSHOP_PORTAL_URL", "https://www.emanualonline.com/cars/").strip() or "https://www.emanualonline.com/cars/"
DEFAULT_LABOUR_RATE = float(os.environ.get("BAM_LABOUR_RATE", "145") or 145)
EMAIL_WEBMAIL_URL = os.environ.get("BAM_EMAIL_WEBMAIL_URL", "").strip()
EMAIL_ADDRESS = os.environ.get("BAM_EMAIL_ADDRESS", "").strip()
EMAIL_IMAP_SERVER = os.environ.get("BAM_EMAIL_IMAP_SERVER", "").strip()
EMAIL_IMAP_PORT = int(os.environ.get("BAM_EMAIL_IMAP_PORT", "993"))
EMAIL_SMTP_SERVER = os.environ.get("BAM_EMAIL_SMTP_SERVER", "").strip()
EMAIL_SMTP_PORT = int(os.environ.get("BAM_EMAIL_SMTP_PORT", "465"))
EMAIL_PASSWORD = os.environ.get("BAM_EMAIL_PASSWORD", "").strip()

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "pdf"}
BACKUP_EXTENSIONS = {"zip"}


@app.context_processor
def inject_app_identity():
    return {"app_version": APP_VERSION, "app_name": APP_NAME}

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def ensure_column(conn, table, column, definition):
    columns = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    if column not in columns:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")



def init_db():
    conn = db()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        display_name TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'owner'
    );

    CREATE TABLE IF NOT EXISTS vehicles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stock_no TEXT UNIQUE NOT NULL,
        status TEXT NOT NULL DEFAULT 'In Stock',
        purchase_date TEXT,
        make TEXT NOT NULL,
        model TEXT NOT NULL,
        variant TEXT,
        year INTEGER,
        vin TEXT UNIQUE,
        registration TEXT,
        odometer_km INTEGER,
        colour TEXT,
        purchase_price_inc_gst REAL NOT NULL DEFAULT 0,
        purchase_gst REAL NOT NULL DEFAULT 0,
        barry_contribution REAL NOT NULL DEFAULT 0,
        matt_contribution REAL NOT NULL DEFAULT 0,
        sale_ownership TEXT NOT NULL DEFAULT 'BAM Joint',
        rego_expiry TEXT,
        photo_filename TEXT,
        notes TEXT,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER NOT NULL,
        expense_date TEXT,
        category TEXT NOT NULL,
        description TEXT NOT NULL,
        supplier TEXT,
        paid_by TEXT NOT NULL,
        cost_inc_gst REAL NOT NULL DEFAULT 0,
        gst_amount REAL NOT NULL DEFAULT 0,
        receipt_filename TEXT,
        notes TEXT,
        FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER UNIQUE NOT NULL,
        sale_date TEXT,
        buyer_name TEXT,
        buyer_phone TEXT,
        sale_price_inc_gst REAL NOT NULL DEFAULT 0,
        sale_gst REAL NOT NULL DEFAULT 0,
        advertising_cost REAL NOT NULL DEFAULT 0,
        transfer_cost REAL NOT NULL DEFAULT 0,
        notes TEXT,
        FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS contacts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contact_type TEXT,
        name TEXT NOT NULL,
        phone TEXT,
        email TEXT,
        address TEXT,
        licence_no TEXT,
        notes TEXT
    );

    CREATE TABLE IF NOT EXISTS consumables (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        purchase_date TEXT,
        item TEXT NOT NULL,
        category TEXT,
        supplier TEXT,
        purchased_by TEXT,
        qty_purchased REAL NOT NULL DEFAULT 0,
        unit_cost_inc_gst REAL NOT NULL DEFAULT 0,
        gst_amount REAL NOT NULL DEFAULT 0,
        qty_used REAL NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS vehicle_photos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER NOT NULL,
        filename TEXT NOT NULL,
        caption TEXT,
        uploaded_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS job_cards (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER NOT NULL,
        job_date TEXT,
        category TEXT,
        description TEXT NOT NULL,
        supplier TEXT,
        paid_by TEXT,
        estimated_cost REAL NOT NULL DEFAULT 0,
        actual_cost_inc_gst REAL NOT NULL DEFAULT 0,
        gst_amount REAL NOT NULL DEFAULT 0,
        status TEXT NOT NULL DEFAULT 'Open',
        notes TEXT,
        FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS job_card_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        job_card_id INTEGER NOT NULL,
        old_status TEXT,
        new_status TEXT NOT NULL,
        changed_by TEXT,
        changed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        note TEXT,
        FOREIGN KEY(job_card_id) REFERENCES job_cards(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS vehicle_documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER NOT NULL,
        document_type TEXT NOT NULL,
        filename TEXT NOT NULL,
        description TEXT,
        uploaded_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS service_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER NOT NULL,
        service_date TEXT,
        odometer_km INTEGER,
        service_type TEXT,
        description TEXT NOT NULL,
        supplier TEXT,
        cost_inc_gst REAL NOT NULL DEFAULT 0,
        gst_amount REAL NOT NULL DEFAULT 0,
        paid_by TEXT,
        next_service_date TEXT,
        notes TEXT,
        FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS parts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        part_number TEXT UNIQUE,
        part_name TEXT NOT NULL,
        category TEXT,
        supplier TEXT,
        quantity_on_hand REAL NOT NULL DEFAULT 0,
        reorder_level REAL NOT NULL DEFAULT 0,
        unit_cost_inc_gst REAL NOT NULL DEFAULT 0,
        gst_amount_per_unit REAL NOT NULL DEFAULT 0,
        storage_location TEXT,
        notes TEXT
    );

    CREATE TABLE IF NOT EXISTS part_usage (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        part_id INTEGER NOT NULL,
        vehicle_id INTEGER NOT NULL,
        job_card_id INTEGER,
        usage_date TEXT,
        quantity_used REAL NOT NULL DEFAULT 0,
        unit_cost_inc_gst REAL NOT NULL DEFAULT 0,
        paid_by TEXT,
        notes TEXT,
        FOREIGN KEY(part_id) REFERENCES parts(id),
        FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE,
        FOREIGN KEY(job_card_id) REFERENCES job_cards(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_name TEXT,
        action TEXT NOT NULL,
        entity_type TEXT,
        entity_id INTEGER,
        details TEXT,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS reminders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER,
        reminder_date TEXT NOT NULL,
        reminder_type TEXT NOT NULL,
        title TEXT NOT NULL,
        notes TEXT,
        completed INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS tasks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER,
        task_date TEXT,
        due_date TEXT,
        title TEXT NOT NULL,
        category TEXT,
        assigned_to TEXT,
        priority TEXT NOT NULL DEFAULT 'Normal',
        status TEXT NOT NULL DEFAULT 'Open',
        notes TEXT,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE
    );
    """)

    # Upgrade columns from older BAM databases.
    ensure_column(conn, "users", "is_active", "INTEGER DEFAULT 1")
    ensure_column(conn, "users", "created_at", "TEXT")
    ensure_column(conn, "users", "last_login", "TEXT")

    ensure_column(conn, "vehicles", "ppsr_number", "TEXT")
    ensure_column(conn, "vehicles", "roadworthy_status", "TEXT DEFAULT 'Not Checked'")
    ensure_column(conn, "vehicles", "service_due_date", "TEXT")
    ensure_column(conn, "vehicles", "service_history", "TEXT")
    ensure_column(conn, "vehicles", "asking_price", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "advertisement_title", "TEXT")
    ensure_column(conn, "vehicles", "advertisement_description", "TEXT")

    ensure_column(conn, "vehicles", "featured_photo_id", "INTEGER")
    ensure_column(conn, "vehicles", "estimated_sale_price", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "minimum_sale_price", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "valuation_notes", "TEXT")

    ensure_column(conn, "vehicles", "advertised_date", "TEXT")
    ensure_column(conn, "vehicles", "ready_for_sale_date", "TEXT")
    ensure_column(conn, "vehicles", "target_profit", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "negotiated_price", "REAL DEFAULT 0")

    ensure_column(conn, "vehicles", "asset_type", "TEXT DEFAULT 'Car'")
    ensure_column(conn, "vehicles", "length_m", "REAL")
    ensure_column(conn, "vehicles", "width_m", "REAL")
    ensure_column(conn, "vehicles", "tare_weight_kg", "REAL")
    ensure_column(conn, "vehicles", "atm_kg", "REAL")
    ensure_column(conn, "vehicles", "gtm_kg", "REAL")
    ensure_column(conn, "vehicles", "berths", "INTEGER")
    ensure_column(conn, "vehicles", "axles", "INTEGER")
    ensure_column(conn, "vehicles", "caravan_features", "TEXT")
    ensure_column(conn, "vehicles", "trailer_features", "TEXT")
    ensure_column(conn, "vehicles", "boat_type", "TEXT")
    ensure_column(conn, "vehicles", "hull_material", "TEXT")
    ensure_column(conn, "vehicles", "engine_make", "TEXT")
    ensure_column(conn, "vehicles", "engine_model", "TEXT")
    ensure_column(conn, "vehicles", "engine_hours", "REAL")
    ensure_column(conn, "vehicles", "horsepower", "REAL")
    ensure_column(conn, "vehicles", "fuel_type", "TEXT")
    ensure_column(conn, "vehicles", "hin", "TEXT")
    ensure_column(conn, "vehicles", "trailer_included", "INTEGER DEFAULT 0")
    ensure_column(conn, "vehicles", "trailer_registration", "TEXT")
    ensure_column(conn, "vehicles", "capacity_people", "INTEGER")
    ensure_column(conn, "vehicles", "boat_features", "TEXT")

    # Version 22 - Dealer Intelligence
    ensure_column(conn, "vehicles", "vin_decode_json", "TEXT")
    ensure_column(conn, "vehicles", "decoded_at", "TEXT")
    ensure_column(conn, "vehicles", "manufacturer_name", "TEXT")
    ensure_column(conn, "vehicles", "body_class", "TEXT")
    ensure_column(conn, "vehicles", "drive_type", "TEXT")
    ensure_column(conn, "vehicles", "transmission_style", "TEXT")
    ensure_column(conn, "vehicles", "engine_cylinders", "TEXT")
    ensure_column(conn, "vehicles", "engine_displacement_l", "TEXT")
    ensure_column(conn, "vehicles", "fuel_type_primary", "TEXT")
    ensure_column(conn, "vehicles", "decoded_model_year", "TEXT")
    ensure_column(conn, "vehicles", "decoded_series", "TEXT")
    ensure_column(conn, "vehicles", "decoded_trim", "TEXT")
    # Version 24.2 - VIN Intelligence & verified workshop specifications
    ensure_column(conn, "vehicles", "vehicle_type_decoded", "TEXT")
    ensure_column(conn, "vehicles", "doors_decoded", "TEXT")
    ensure_column(conn, "vehicles", "engine_model_decoded", "TEXT")
    ensure_column(conn, "vehicles", "engine_power_kw_decoded", "TEXT")
    ensure_column(conn, "vehicles", "plant_country_decoded", "TEXT")
    ensure_column(conn, "vehicles", "brake_system_decoded", "TEXT")
    ensure_column(conn, "vehicles", "turbo_decoded", "TEXT")
    ensure_column(conn, "vehicles", "electrification_level_decoded", "TEXT")
    ensure_column(conn, "vehicles", "paint_code", "TEXT")
    ensure_column(conn, "vehicles", "engine_code", "TEXT")
    ensure_column(conn, "vehicles", "transmission_code", "TEXT")
    ensure_column(conn, "vehicles", "option_codes", "TEXT")
    ensure_column(conn, "vehicles", "service_interval_km", "INTEGER")
    ensure_column(conn, "vehicles", "service_interval_months", "INTEGER")
    ensure_column(conn, "vehicles", "technical_data_source", "TEXT")
    ensure_column(conn, "vehicles", "technical_notes", "TEXT")
    ensure_column(conn, "vehicles", "market_price_low", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "market_price_mid", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "market_price_high", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "market_price_checked_at", "TEXT")
    ensure_column(conn, "vehicles", "comparable_price_1", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "comparable_price_2", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "comparable_price_3", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "comparable_price_4", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "comparable_price_5", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "private_value_low", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "private_value_high", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "wholesale_value_low", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "wholesale_value_high", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "trade_value_low", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "trade_value_high", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "dealer_value_low", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "dealer_value_high", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "valuation_provider", "TEXT")
    ensure_column(conn, "vehicles", "valuation_confidence", "TEXT")
    ensure_column(conn, "vehicles", "valuation_source", "TEXT")

    # Version 18 - Parts Vehicle / Dismantling
    ensure_column(conn, "vehicles", "vehicle_purpose", "TEXT DEFAULT 'Retail Sale'")
    ensure_column(conn, "vehicles", "dismantling_status", "TEXT DEFAULT 'Not Started'")
    ensure_column(conn, "vehicles", "shell_sale_price", "REAL DEFAULT 0")
    ensure_column(conn, "vehicles", "shell_sale_date", "TEXT")

    ensure_column(conn, "sales", "invoice_number", "TEXT")
    ensure_column(conn, "sales", "buyer_email", "TEXT")
    ensure_column(conn, "sales", "buyer_address", "TEXT")

    ensure_column(conn, "sales", "deposit_amount", "REAL DEFAULT 0")
    ensure_column(conn, "sales", "deposit_date", "TEXT")
    ensure_column(conn, "sales", "payment_method", "TEXT")
    ensure_column(conn, "sales", "trade_in_description", "TEXT")
    ensure_column(conn, "sales", "trade_in_value", "REAL DEFAULT 0")
    ensure_column(conn, "sales", "warranty_type", "TEXT")
    ensure_column(conn, "sales", "warranty_expiry", "TEXT")
    ensure_column(conn, "sales", "contract_number", "TEXT")
    ensure_column(conn, "sales", "invoice_status", "TEXT DEFAULT 'Draft'")
    ensure_column(conn, "sales", "updated_at", "TEXT")
    ensure_column(conn, "sales", "delivery_status", "TEXT DEFAULT 'Preparing'")
    ensure_column(conn, "sales", "delivery_date", "TEXT")
    ensure_column(conn, "sales", "keys_handed_over", "INTEGER DEFAULT 0")
    ensure_column(conn, "sales", "registration_transferred", "INTEGER DEFAULT 0")
    ensure_column(conn, "sales", "customer_signature_received", "INTEGER DEFAULT 0")
    ensure_column(conn, "sales", "finance_documents_complete", "INTEGER DEFAULT 0")
    ensure_column(conn, "sales", "warranty_documents_complete", "INTEGER DEFAULT 0")
    ensure_column(conn, "sales", "deal_notes", "TEXT")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS invoice_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sale_id INTEGER NOT NULL,
            vehicle_id INTEGER NOT NULL,
            invoice_number TEXT,
            changed_by TEXT,
            change_note TEXT,
            old_values TEXT,
            new_values TEXT,
            changed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(sale_id) REFERENCES sales(id) ON DELETE CASCADE,
            FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE
        )
    """)

    conn.executescript("""
        CREATE TABLE IF NOT EXISTS part_photos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            caption TEXT,
            is_featured INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (part_id)
                REFERENCES parts(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS part_sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_id INTEGER NOT NULL,
            quantity INTEGER DEFAULT 1,
            customer_name TEXT,
            customer_phone TEXT,
            customer_email TEXT,
            sale_price REAL DEFAULT 0,
            freight_cost REAL DEFAULT 0,
            payment_method TEXT,
            warranty TEXT,
            invoice_number TEXT,
            sale_date TEXT DEFAULT CURRENT_TIMESTAMP,
            notes TEXT,
            FOREIGN KEY (part_id) REFERENCES parts(id)
        );


        CREATE TABLE IF NOT EXISTS part_shipments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shipment_number TEXT UNIQUE NOT NULL,
            part_id INTEGER,
            quantity REAL NOT NULL DEFAULT 1,
            customer_name TEXT NOT NULL,
            customer_phone TEXT,
            customer_email TEXT,
            address_line TEXT NOT NULL,
            suburb TEXT,
            state TEXT,
            postcode TEXT,
            courier TEXT,
            tracking_number TEXT,
            parcel_weight_kg REAL DEFAULT 0,
            parcel_length_cm REAL DEFAULT 0,
            parcel_width_cm REAL DEFAULT 0,
            parcel_height_cm REAL DEFAULT 0,
            shipping_cost REAL DEFAULT 0,
            freight_charged REAL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'Ready to Pack',
            date_created TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            date_sent TEXT,
            date_delivered TEXT,
            invoice_included INTEGER NOT NULL DEFAULT 0,
            bubble_wrapped INTEGER NOT NULL DEFAULT 0,
            box_sealed INTEGER NOT NULL DEFAULT 0,
            tracking_sent INTEGER NOT NULL DEFAULT 0,
            stock_adjusted INTEGER NOT NULL DEFAULT 0,
            notes TEXT,
            FOREIGN KEY(part_id) REFERENCES parts(id) ON DELETE SET NULL
        );
    """)
    ensure_column(conn, "parts", "vehicle_id", "INTEGER")
    ensure_column(conn, "parts", "vehicle_stock_no", "TEXT")
    ensure_column(conn, "parts", "vin", "TEXT")
    ensure_column(conn, "parts", "make", "TEXT")
    ensure_column(conn, "parts", "model", "TEXT")
    ensure_column(conn, "parts", "year", "INTEGER")
    ensure_column(conn, "parts", "subcategory", "TEXT")
    ensure_column(conn, "parts", "description", "TEXT")
    ensure_column(conn, "parts", "condition", "TEXT DEFAULT 'Used'")
    ensure_column(conn, "parts", "selling_price", "REAL DEFAULT 0")
    ensure_column(conn, "parts", "status", "TEXT DEFAULT 'In Stock'")
    ensure_column(conn, "parts", "engine_code", "TEXT")
    ensure_column(conn, "parts", "transmission_code", "TEXT")
    ensure_column(conn, "parts", "barcode", "TEXT")
    ensure_column(conn, "parts", "date_added", "TEXT")

    # Version 19 - Professional Parts Inventory
    ensure_column(conn, "parts", "position", "TEXT")
    ensure_column(conn, "parts", "fitment", "TEXT")
    ensure_column(conn, "parts", "manufacturer_part_no", "TEXT")
    ensure_column(conn, "parts", "reserved_for", "TEXT")
    ensure_column(conn, "parts", "reserved_until", "TEXT")
    ensure_column(conn, "parts", "updated_at", "TEXT")

    # Version 24.3 - Parts Intelligence + QR / Barcode workflow
    ensure_column(conn, "parts", "alternate_part_numbers", "TEXT")
    ensure_column(conn, "parts", "interchange_notes", "TEXT")
    ensure_column(conn, "parts", "warranty_days", "INTEGER DEFAULT 0")
    ensure_column(conn, "parts", "weight_kg", "REAL DEFAULT 0")
    ensure_column(conn, "parts", "length_cm", "REAL DEFAULT 0")
    ensure_column(conn, "parts", "width_cm", "REAL DEFAULT 0")
    ensure_column(conn, "parts", "height_cm", "REAL DEFAULT 0")
    ensure_column(conn, "parts", "inventory_last_checked", "TEXT")
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS part_location_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_id INTEGER NOT NULL,
            old_location TEXT,
            new_location TEXT,
            moved_by TEXT,
            notes TEXT,
            moved_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(part_id) REFERENCES parts(id) ON DELETE CASCADE
        );
    """)
    conn.execute("UPDATE parts SET barcode=part_number WHERE COALESCE(barcode,'')='' AND COALESCE(part_number,'')!=''")
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS equipment (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            equipment_no TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            category TEXT,
            brand TEXT,
            model TEXT,
            serial_number TEXT,
            purchase_date TEXT,
            purchase_price REAL NOT NULL DEFAULT 0,
            current_value REAL NOT NULL DEFAULT 0,
            supplier TEXT,
            warranty_expiry TEXT,
            location TEXT,
            assigned_to TEXT,
            condition TEXT NOT NULL DEFAULT 'Good',
            status TEXT NOT NULL DEFAULT 'Available',
            next_service_date TEXT,
            calibration_due TEXT,
            test_tag_due TEXT,
            photo_filename TEXT,
            receipt_filename TEXT,
            notes TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS equipment_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            equipment_id INTEGER NOT NULL,
            action TEXT NOT NULL,
            assigned_to TEXT,
            location TEXT,
            condition TEXT,
            notes TEXT,
            action_date TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            user_name TEXT,
            FOREIGN KEY(equipment_id) REFERENCES equipment(id) ON DELETE CASCADE
        );
    """)

    # Version 23 - Workshop Intelligence
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS workshop_operations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vehicle_id INTEGER,
            part_id INTEGER,
            operation_code TEXT,
            system_name TEXT,
            operation_name TEXT NOT NULL,
            labour_hours REAL NOT NULL DEFAULT 0,
            labour_rate REAL NOT NULL DEFAULT 0,
            source_name TEXT,
            source_reference TEXT,
            procedure_url TEXT,
            notes TEXT,
            make TEXT,
            model TEXT,
            year INTEGER,
            engine_hint TEXT,
            created_by TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE,
            FOREIGN KEY(part_id) REFERENCES parts(id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS workshop_references (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vehicle_id INTEGER NOT NULL,
            reference_type TEXT NOT NULL,
            title TEXT NOT NULL,
            provider_name TEXT,
            reference_url TEXT,
            reference_code TEXT,
            notes TEXT,
            created_by TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE CASCADE
        );
    """)
    ensure_column(conn, "job_cards", "labour_operation_id", "INTEGER")
    ensure_column(conn, "job_cards", "labour_hours", "REAL DEFAULT 0")
    ensure_column(conn, "job_cards", "labour_rate", "REAL DEFAULT 0")
    ensure_column(conn, "job_cards", "labour_source", "TEXT")
    ensure_column(conn, "job_cards", "labour_code", "TEXT")
    ensure_column(conn, "job_cards", "procedure_url", "TEXT")

    # Version 23.1 - eManualOnline + Email Centre
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS email_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            direction TEXT NOT NULL DEFAULT 'Incoming',
            message_date TEXT,
            sender TEXT,
            recipient TEXT,
            subject TEXT NOT NULL,
            body_excerpt TEXT,
            status TEXT NOT NULL DEFAULT 'Unread',
            follow_up_date TEXT,
            vehicle_id INTEGER,
            contact_id INTEGER,
            external_reference TEXT,
            created_by TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE SET NULL,
            FOREIGN KEY(contact_id) REFERENCES contacts(id) ON DELETE SET NULL
        );
    """)

    # Version 24.4 - CRM + Communications Centre
    ensure_column(conn, "contacts", "company", "TEXT")
    ensure_column(conn, "contacts", "crm_status", "TEXT DEFAULT 'Active'")
    ensure_column(conn, "contacts", "source", "TEXT")
    ensure_column(conn, "contacts", "tags", "TEXT")
    ensure_column(conn, "contacts", "preferred_contact", "TEXT")
    ensure_column(conn, "contacts", "last_contact_date", "TEXT")
    ensure_column(conn, "contacts", "next_follow_up_date", "TEXT")
    ensure_column(conn, "contacts", "updated_at", "TEXT")
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS crm_activities (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            contact_id INTEGER NOT NULL,
            vehicle_id INTEGER,
            email_message_id INTEGER,
            activity_date TEXT NOT NULL,
            activity_type TEXT NOT NULL,
            subject TEXT NOT NULL,
            notes TEXT,
            outcome TEXT,
            follow_up_date TEXT,
            created_by TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(contact_id) REFERENCES contacts(id) ON DELETE CASCADE,
            FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE SET NULL,
            FOREIGN KEY(email_message_id) REFERENCES email_messages(id) ON DELETE SET NULL
        );
        CREATE INDEX IF NOT EXISTS idx_crm_activity_contact ON crm_activities(contact_id);
        CREATE INDEX IF NOT EXISTS idx_crm_activity_followup ON crm_activities(follow_up_date);
    """)
    ensure_column(conn, "email_messages", "priority", "TEXT DEFAULT 'Normal'")
    ensure_column(conn, "email_messages", "category", "TEXT")
    ensure_column(conn, "email_messages", "assigned_to", "TEXT")
    ensure_column(conn, "vehicles", "sale_ownership", "TEXT NOT NULL DEFAULT 'BAM Joint'")
    # Version 24.1 - Professional Workshop Scheduler
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS workshop_bookings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vehicle_id INTEGER,
            contact_id INTEGER,
            booking_date TEXT NOT NULL,
            start_time TEXT,
            end_time TEXT,
            technician TEXT,
            bay TEXT,
            job_type TEXT,
            description TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Booked',
            quoted_hours REAL NOT NULL DEFAULT 0,
            labour_rate REAL NOT NULL DEFAULT 0,
            customer_name TEXT,
            customer_phone TEXT,
            notes TEXT,
            job_card_id INTEGER,
            created_by TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT,
            FOREIGN KEY(vehicle_id) REFERENCES vehicles(id) ON DELETE SET NULL,
            FOREIGN KEY(contact_id) REFERENCES contacts(id) ON DELETE SET NULL,
            FOREIGN KEY(job_card_id) REFERENCES job_cards(id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS workshop_time_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            booking_id INTEGER NOT NULL,
            technician TEXT,
            clock_in TEXT NOT NULL,
            clock_out TEXT,
            minutes INTEGER NOT NULL DEFAULT 0,
            notes TEXT,
            created_by TEXT,
            FOREIGN KEY(booking_id) REFERENCES workshop_bookings(id) ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_workshop_bookings_date ON workshop_bookings(booking_date);
        CREATE INDEX IF NOT EXISTS idx_workshop_bookings_vehicle ON workshop_bookings(vehicle_id);
        CREATE INDEX IF NOT EXISTS idx_workshop_time_booking ON workshop_time_entries(booking_id);
    """)

    # Version 25.1 - BAM Auction Watch
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS auction_vehicles (
            id INTEGER PRIMARY KEY AUTOINCREMENT, status TEXT NOT NULL DEFAULT 'Watching', asset_type TEXT NOT NULL DEFAULT 'Car',
            auction_name TEXT, auction_location TEXT, auction_url TEXT, lot_number TEXT, auction_start TEXT, auction_finish TEXT,
            year INTEGER, make TEXT NOT NULL, model TEXT NOT NULL, variant TEXT, vin TEXT, registration TEXT, odometer_km INTEGER,
            engine_hours REAL, colour TEXT, interior TEXT, transmission TEXT, drive_type TEXT, fuel_type TEXT, tow_bar INTEGER NOT NULL DEFAULT 0,
            condition_grade TEXT, condition_notes TEXT, current_bid REAL NOT NULL DEFAULT 0, max_bid REAL NOT NULL DEFAULT 0,
            sold_price REAL NOT NULL DEFAULT 0, auction_fees REAL NOT NULL DEFAULT 0, transport_cost REAL NOT NULL DEFAULT 0, other_costs REAL NOT NULL DEFAULT 0,
            market_low REAL NOT NULL DEFAULT 0, market_mid REAL NOT NULL DEFAULT 0, market_high REAL NOT NULL DEFAULT 0, valuation_source TEXT, valuation_checked_at TEXT,
            won_vehicle_id INTEGER, created_by TEXT, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, updated_at TEXT,
            FOREIGN KEY(won_vehicle_id) REFERENCES vehicles(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS auction_photos (id INTEGER PRIMARY KEY AUTOINCREMENT, auction_vehicle_id INTEGER NOT NULL, filename TEXT NOT NULL, caption TEXT, uploaded_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY(auction_vehicle_id) REFERENCES auction_vehicles(id) ON DELETE CASCADE);
        CREATE INDEX IF NOT EXISTS idx_auction_vehicle_type ON auction_vehicles(asset_type);
        CREATE INDEX IF NOT EXISTS idx_auction_make_model ON auction_vehicles(make,model);
        CREATE INDEX IF NOT EXISTS idx_auction_finish ON auction_vehicles(auction_finish);
        CREATE INDEX IF NOT EXISTS idx_auction_status ON auction_vehicles(status);
    """)

    # Version 25.5 - richer Auction Watch fields for cars, caravans, boats,
    # trailers and motorcycles. ensure_column keeps existing databases safe.
    ensure_column(conn, "auction_vehicles", "engine_size", "TEXT")
    ensure_column(conn, "auction_vehicles", "engine_cc", "INTEGER")
    ensure_column(conn, "auction_vehicles", "engine_cylinders", "TEXT")
    ensure_column(conn, "auction_vehicles", "length_m", "REAL")
    ensure_column(conn, "auction_vehicles", "berths", "INTEGER")
    ensure_column(conn, "auction_vehicles", "axles", "INTEGER")
    ensure_column(conn, "auction_vehicles", "tare_weight_kg", "REAL")
    ensure_column(conn, "auction_vehicles", "atm_kg", "REAL")
    ensure_column(conn, "auction_vehicles", "gtm_kg", "REAL")
    ensure_column(conn, "auction_vehicles", "ball_weight_kg", "REAL")
    ensure_column(conn, "auction_vehicles", "width_m", "REAL")
    ensure_column(conn, "auction_vehicles", "height_m", "REAL")
    ensure_column(conn, "auction_vehicles", "caravan_features", "TEXT")
    ensure_column(conn, "auction_vehicles", "boat_type", "TEXT")
    ensure_column(conn, "auction_vehicles", "hull_material", "TEXT")
    ensure_column(conn, "auction_vehicles", "engine_make", "TEXT")
    ensure_column(conn, "auction_vehicles", "engine_model", "TEXT")
    ensure_column(conn, "auction_vehicles", "horsepower", "REAL")
    ensure_column(conn, "auction_vehicles", "trailer_included", "INTEGER DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "trailer_registration", "TEXT")
    ensure_column(conn, "auction_vehicles", "capacity_people", "INTEGER")
    ensure_column(conn, "auction_vehicles", "boat_features", "TEXT")
    ensure_column(conn, "auction_vehicles", "trailer_features", "TEXT")

    # Version 25.6 - BAM Buying Watch + unified valuation hub.
    ensure_column(conn, "auction_vehicles", "listing_source", "TEXT DEFAULT 'Auction'")
    ensure_column(conn, "auction_vehicles", "seller_name", "TEXT")
    ensure_column(conn, "auction_vehicles", "seller_phone", "TEXT")
    ensure_column(conn, "auction_vehicles", "seller_location", "TEXT")
    ensure_column(conn, "auction_vehicles", "listing_url", "TEXT")
    ensure_column(conn, "auction_vehicles", "date_first_seen", "TEXT")
    ensure_column(conn, "auction_vehicles", "last_checked", "TEXT")
    ensure_column(conn, "auction_vehicles", "asking_price", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "negotiated_price", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "private_value_low", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "private_value_high", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "wholesale_value_low", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "wholesale_value_high", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "trade_value_low", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "trade_value_high", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "dealer_value_low", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "dealer_value_high", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "suggested_buy_price", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "valuation_provider", "TEXT")
    ensure_column(conn, "auction_vehicles", "valuation_confidence", "TEXT")
    # Version 25.8.3 - registration sale/status from auction listings.
    ensure_column(conn, "auction_vehicles", "registration_status", "TEXT")
    # Version 25.9 - richer Buying Watch vehicle details.
    ensure_column(conn, "auction_vehicles", "reserve_status", "TEXT")
    ensure_column(conn, "auction_vehicles", "body_type", "TEXT")
    ensure_column(conn, "auction_vehicles", "seat_count", "INTEGER")
    # Version 25.14 - Auction Buy Calculator & Deal Score.
    ensure_column(conn, "auction_vehicles", "quick_sale_value", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "repair_allowance", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "target_profit", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "rego_ppsr_cost", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "boat_engine_cost", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "boat_hull_cost", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "boat_trailer_cost", "REAL DEFAULT 0")
    # Version 25.16 - user-entered Australian market comparables.
    ensure_column(conn, "auction_vehicles", "comparable_price_1", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "comparable_price_2", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "comparable_price_3", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "comparable_price_4", "REAL DEFAULT 0")
    ensure_column(conn, "auction_vehicles", "comparable_price_5", "REAL DEFAULT 0")

    count = conn.execute(
        "SELECT COUNT(*) AS c FROM users"
    ).fetchone()["c"]

    if count == 0:
        conn.execute(
            "INSERT INTO users(username,password_hash,display_name,role) VALUES(?,?,?,?)",
            ("barry", generate_password_hash("ChangeMe123!"), "Barry", "owner"),
        )
        conn.execute(
            "INSERT INTO users(username,password_hash,display_name,role) VALUES(?,?,?,?)",
            ("matt", generate_password_hash("ChangeMe123!"), "Matt", "owner"),
        )

    conn.commit()
    conn.close()



def next_stock_number(conn=None):
    own_conn = conn is None
    conn = conn or db()

    highest = 0
    rows = conn.execute(
        "SELECT stock_no FROM vehicles WHERE stock_no LIKE 'BAM-%'"
    ).fetchall()

    for row in rows:
        match = re.search(r"(\d+)$", row["stock_no"] or "")
        if match:
            highest = max(highest, int(match.group(1)))

    if own_conn:
        conn.close()

    return f"BAM-{highest + 1:05d}"


def next_invoice_number(conn=None):
    own_conn = conn is None
    conn = conn or db()

    highest = 0
    rows = conn.execute(
        "SELECT invoice_number FROM sales WHERE invoice_number LIKE 'INV-%'"
    ).fetchall()

    for row in rows:
        match = re.search(r"(\d+)$", row["invoice_number"] or "")
        if match:
            highest = max(highest, int(match.group(1)))

    if own_conn:
        conn.close()

    return f"INV-{highest + 1:05d}"



def next_part_number(conn=None):
    own_conn = conn is None
    conn = conn or db()
    highest = 0
    rows = conn.execute("SELECT part_number FROM parts WHERE part_number LIKE 'PRT-%'").fetchall()
    for row in rows:
        match = re.search(r"(\d+)$", row["part_number"] or "")
        if match:
            highest = max(highest, int(match.group(1)))
    if own_conn:
        conn.close()
    return f"PRT-{highest + 1:06d}"


def next_shipment_number(conn=None):
    own_conn = conn is None
    conn = conn or db()
    highest = 0
    rows = conn.execute(
        "SELECT shipment_number FROM part_shipments WHERE shipment_number LIKE 'SHP-%'"
    ).fetchall()
    for row in rows:
        match = re.search(r"(\d+)$", row["shipment_number"] or "")
        if match:
            highest = max(highest, int(match.group(1)))
    if own_conn:
        conn.close()
    return f"SHP-{highest + 1:05d}"

def local_network_ip():
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.connect(("8.8.8.8", 80))
        ip = sock.getsockname()[0]
        sock.close()
        return ip
    except OSError:
        return "127.0.0.1"


def log_action(action, entity_type=None, entity_id=None, details=None):
    conn = db()
    conn.execute("""
        INSERT INTO audit_log(user_name,action,entity_type,entity_id,details)
        VALUES(?,?,?,?,?)
    """, (
        session.get("display_name"),
        action,
        entity_type,
        entity_id,
        details,
    ))
    conn.commit()
    conn.close()


def next_contract_number(conn=None):
    own_conn = conn is None
    conn = conn or db()

    highest = 0
    rows = conn.execute(
        "SELECT contract_number FROM sales WHERE contract_number LIKE 'CON-%'"
    ).fetchall()

    for row in rows:
        match = re.search(r"(\d+)$", row["contract_number"] or "")
        if match:
            highest = max(highest, int(match.group(1)))

    if own_conn:
        conn.close()

    return f"CON-{highest + 1:05d}"



def _clean_text(value):
    value = str(value or "").strip()
    return value if value and value.lower() not in {"not applicable", "not available", "null", "none"} else ""


def decode_vin(vin):
    """Decode a VIN using a configurable external decoder. Returns (data, error)."""
    vin = re.sub(r"[^A-Za-z0-9]", "", vin or "").upper()
    if len(vin) != 17:
        return None, "VIN must contain exactly 17 characters."
    try:
        url = VIN_DECODER_URL.format(vin=urllib.parse.quote(vin))
        req = urllib.request.Request(url, headers={"User-Agent": f"{APP_NAME}/{APP_VERSION}"})
        with urllib.request.urlopen(req, timeout=12) as response:
            payload = json.loads(response.read().decode("utf-8"))
        result = (payload.get("Results") or [{}])[0]
        if not result:
            return None, "VIN decoder returned no vehicle data."
        specs = {
            "vin": vin,
            "make": _clean_text(result.get("Make")),
            "model": _clean_text(result.get("Model")),
            "year": _clean_text(result.get("ModelYear")),
            "manufacturer_name": _clean_text(result.get("Manufacturer")),
            "body_class": _clean_text(result.get("BodyClass")),
            "drive_type": _clean_text(result.get("DriveType")),
            "transmission_style": _clean_text(result.get("TransmissionStyle")),
            "engine_cylinders": _clean_text(result.get("EngineCylinders")),
            "engine_displacement_l": _clean_text(result.get("DisplacementL")),
            "fuel_type_primary": _clean_text(result.get("FuelTypePrimary")),
            "series": _clean_text(result.get("Series")),
            "trim": _clean_text(result.get("Trim")),
            "vehicle_type": _clean_text(result.get("VehicleType")),
            "doors": _clean_text(result.get("Doors")),
            "engine_model": _clean_text(result.get("EngineModel")),
            "engine_power_kw": _clean_text(result.get("EngineKW")),
            "plant_country": _clean_text(result.get("PlantCountry")),
            "brake_system": _clean_text(result.get("BrakeSystemType")),
            "turbo": _clean_text(result.get("Turbo")),
            "electrification_level": _clean_text(result.get("ElectrificationLevel")),
            "raw": result,
        }
        error_text = _clean_text(result.get("ErrorText"))
        if not any(specs.get(k) for k in ("make", "model", "year", "body_class")):
            return specs, error_text or "VIN decoded, but the provider returned limited specifications."
        return specs, None
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, ValueError, json.JSONDecodeError) as exc:
        return None, f"VIN service unavailable: {exc}"


def duplicate_vehicle_matches(conn, vin="", registration="", stock_no="", exclude_id=None):
    checks = []
    params = []
    vin = (vin or "").strip().upper()
    registration = (registration or "").strip().upper()
    stock_no = (stock_no or "").strip().upper()
    if vin:
        checks.append("UPPER(COALESCE(vin,''))=?")
        params.append(vin)
    if registration:
        checks.append("UPPER(COALESCE(registration,''))=?")
        params.append(registration)
    if stock_no:
        checks.append("UPPER(COALESCE(stock_no,''))=?")
        params.append(stock_no)
    if not checks:
        return []
    sql = "SELECT id,stock_no,year,make,model,vin,registration,status FROM vehicles WHERE (" + " OR ".join(checks) + ")"
    if exclude_id:
        sql += " AND id<>?"
        params.append(exclude_id)
    return conn.execute(sql, params).fetchall()


def market_price_suggestion(conn, vehicle):
    rows = conn.execute("""
        SELECT s.sale_price_inc_gst AS price, v.year, v.make, v.model
        FROM sales s JOIN vehicles v ON v.id=s.vehicle_id
        WHERE s.sale_price_inc_gst>0
          AND LOWER(v.make)=LOWER(?) AND LOWER(v.model)=LOWER(?)
          AND v.id<>?
        ORDER BY s.sale_date DESC
        LIMIT 20
    """, (vehicle["make"], vehicle["model"], vehicle["id"])).fetchall()
    source = "same make/model sales in BAM"
    if not rows:
        rows = conn.execute("""
            SELECT s.sale_price_inc_gst AS price, v.year, v.make, v.model
            FROM sales s JOIN vehicles v ON v.id=s.vehicle_id
            WHERE s.sale_price_inc_gst>0 AND LOWER(v.make)=LOWER(?) AND v.id<>?
            ORDER BY s.sale_date DESC LIMIT 20
        """, (vehicle["make"], vehicle["id"])).fetchall()
        source = "same-make sales in BAM"
    prices = sorted(float(r["price"] or 0) for r in rows if float(r["price"] or 0)>0)
    floor = max(float(vehicle["purchase_price_inc_gst"] or 0), float(vehicle["minimum_sale_price"] or 0))
    if prices:
        mid = statistics.median(prices)
        low = prices[max(0, int((len(prices)-1)*0.25))]
        high = prices[min(len(prices)-1, int((len(prices)-1)*0.75))]
        if floor and mid < floor:
            mid = floor * 1.08
            low = floor
            high = max(high, mid * 1.08)
        confidence = "Good" if len(prices) >= 5 else "Limited"
    else:
        base = floor or float(vehicle["asking_price"] or 0) or 0
        low, mid, high = base, base * 1.10 if base else 0, base * 1.20 if base else 0
        source = "vehicle cost/asking price (no comparable BAM sales yet)"
        confidence = "Estimate only"
    return {"low": round(low,2), "mid": round(mid,2), "high": round(high,2), "count": len(prices), "source": source, "confidence": confidence}


def parts_price_suggestion(conn, part):
    rows = conn.execute("""
        SELECT ps.sale_price, ps.quantity
        FROM part_sales ps JOIN parts p ON p.id=ps.part_id
        WHERE ps.sale_price>0 AND (LOWER(p.part_name)=LOWER(?) OR (COALESCE(?, '')<>'' AND LOWER(COALESCE(p.category,''))=LOWER(?)))
        ORDER BY ps.sale_date DESC LIMIT 30
    """, (part["part_name"], part["category"], part["category"])).fetchall()
    unit_prices = [float(r["sale_price"] or 0)/max(float(r["quantity"] or 1), 1) for r in rows if float(r["sale_price"] or 0)>0]
    if unit_prices:
        mid = statistics.median(unit_prices)
        return {"suggested": round(mid,2), "low": round(min(unit_prices),2), "high": round(max(unit_prices),2), "count": len(unit_prices), "source": "recorded BAM parts sales"}
    current = float(part["selling_price"] or 0)
    cost = float(part["unit_cost_inc_gst"] or 0)
    suggestion = current or (cost * 1.5 if cost else 0)
    return {"suggested": round(suggestion,2), "low": round(suggestion*0.9,2) if suggestion else 0, "high": round(suggestion*1.15,2) if suggestion else 0, "count": 0, "source": "current price/cost (no comparable sales yet)"}


def generate_vehicle_ad_text(vehicle):
    title = f"{vehicle['year'] or ''} {vehicle['make']} {vehicle['model']} {vehicle['variant'] or ''}".strip()
    facts = []
    if vehicle["odometer_km"]:
        facts.append(f"{int(vehicle['odometer_km']):,} km")
    if vehicle["colour"]:
        facts.append(str(vehicle["colour"]))
    if vehicle["fuel_type_primary"] or vehicle["fuel_type"]:
        facts.append(str(vehicle["fuel_type_primary"] or vehicle["fuel_type"]))
    fallback = title + "\n\n" + " • ".join(facts) + ("\n\n" + str(vehicle["notes"]) if vehicle["notes"] else "") + "\n\nContact BAM Motor Group — Buy • Sell • Trade."
    if not OPENAI_API_KEY:
        return fallback, "Professional local draft (set OPENAI_API_KEY in Azure to enable AI generation)."
    prompt = (
        "Write a clear, professional Australian used-vehicle advertisement. Do not invent features. "
        "Use only the supplied facts. Include a short headline, a concise paragraph and bullet points. "
        "Avoid exaggerated claims. Facts: " + json.dumps({
            "stock_no": vehicle["stock_no"], "year": vehicle["year"], "make": vehicle["make"],
            "model": vehicle["model"], "variant": vehicle["variant"], "odometer_km": vehicle["odometer_km"],
            "colour": vehicle["colour"], "registration": vehicle["registration"], "roadworthy": vehicle["roadworthy_status"],
            "body_class": vehicle["body_class"], "drive_type": vehicle["drive_type"], "fuel": vehicle["fuel_type_primary"] or vehicle["fuel_type"],
            "asking_price": vehicle["asking_price"], "notes": vehicle["notes"],
        }, ensure_ascii=False)
    )
    try:
        payload = json.dumps({"model": OPENAI_MODEL, "input": prompt, "store": False}).encode("utf-8")
        req = urllib.request.Request("https://api.openai.com/v1/responses", data=payload, headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"}, method="POST")
        with urllib.request.urlopen(req, timeout=35) as response:
            data = json.loads(response.read().decode("utf-8"))
        text = data.get("output_text")
        if not text:
            parts = []
            for item in data.get("output", []):
                for content in item.get("content", []):
                    if content.get("type") == "output_text" and content.get("text"):
                        parts.append(content["text"])
            text = "\n".join(parts).strip()
        return text or fallback, f"AI draft using {OPENAI_MODEL}"
    except Exception as exc:
        return fallback, f"AI service unavailable; local draft used ({exc})."



def _extract_response_text(data):
    text = data.get("output_text") or ""
    if text:
        return text.strip()
    parts = []
    for item in data.get("output", []):
        if item.get("type") == "message":
            for content in item.get("content", []):
                if content.get("type") == "output_text" and content.get("text"):
                    parts.append(content["text"])
    return "\n".join(parts).strip()


def bam_live_market_valuation(details):
    """Research current Australian advertised comparables with OpenAI web search."""
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY is not configured in Azure.")
    clean = {k: v for k, v in dict(details).items() if v not in (None, "", 0, 0.0, "0")}
    asset = str(clean.get("asset_type") or "Car")
    identity = " ".join(str(clean.get(k) or "") for k in ("year", "make", "model", "variant", "auction_name", "stock_no")).strip()
    if not identity:
        raise ValueError("Enter or import identifying details before running BAM Automatic Market Valuation.")
    prompt = f"""You are the market-research engine inside BAM Motor Group, Australia.
Research CURRENT Australian advertised listings for the supplied {asset}. Use web search. Prefer close matches in year, make, model, variant/specification, kilometres/hours, size and equipment. Do not invent listings or prices. Ignore obviously unrelated items and auction guide prices when better retail/private comparables exist.

Vehicle/asset details:\n{json.dumps(clean, ensure_ascii=False)}

Return ONLY valid JSON with this exact shape:
{{"comparables":[{{"price":18500,"title":"short listing title","source":"site/domain","url":"https://..."}}],"confidence":"Good|Limited|Low","summary":"short factual explanation"}}
Use AUD asking prices as numbers only. Return up to 5 genuine useful comparables. If fewer than 2 credible priced listings are found, return whatever genuine comparables exist and set confidence Low. Never manufacture missing prices."""
    payload = json.dumps({
        "model": OPENAI_MODEL,
        "input": prompt,
        "tools": [{"type": "web_search_preview", "search_context_size": "medium"}],
        "tool_choice": "auto",
        "store": False,
    }).encode("utf-8")
    req = urllib.request.Request("https://api.openai.com/v1/responses", data=payload, headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=75) as response:
            data = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")[:800]
        raise RuntimeError(f"OpenAI valuation request failed ({exc.code}): {body}") from exc
    text = _extract_response_text(data)
    match = re.search(r"\{.*\}", text, re.S)
    if not match:
        raise RuntimeError("The live valuation service did not return usable market data.")
    try:
        result = json.loads(match.group(0))
    except json.JSONDecodeError as exc:
        raise RuntimeError("The live valuation service returned invalid market data.") from exc
    comps = []
    for c in result.get("comparables", [])[:5]:
        try:
            price = float(str(c.get("price", 0)).replace("$", "").replace(",", ""))
        except (TypeError, ValueError):
            price = 0
        if price > 0:
            comps.append({"price": round(price, 2), "title": str(c.get("title") or "")[:180], "source": str(c.get("source") or "")[:120], "url": str(c.get("url") or "")[:500]})
    if not comps:
        raise RuntimeError("No credible priced Australian comparables were found. Try adding more identifying details.")
    prices = sorted(c["price"] for c in comps)
    low, high, mid = prices[0], prices[-1], statistics.median(prices)
    return {
        "comparables": comps,
        "prices": [c["price"] for c in comps],
        "market_low": round(low,2), "market_mid": round(mid,2), "market_high": round(high,2),
        "private_low": round(low*.92,2), "private_high": round(high*.97,2),
        "wholesale_low": round(low*.62,2), "wholesale_high": round(mid*.72,2),
        "trade_low": round(low*.68,2), "trade_high": round(mid*.78,2),
        "dealer_low": round(low,2), "dealer_high": round(high,2),
        "confidence": str(result.get("confidence") or ("Good" if len(prices)>=4 else "Limited" if len(prices)>=2 else "Low"))[:40],
        "summary": str(result.get("summary") or "Current Australian advertised comparable listings researched by BAM AI web search.")[:1000],
    }

def ensure_smart_reminders(conn):
    today = date.today()
    def add_once(vehicle_id, due_date, title, notes):
        exists = conn.execute("SELECT id FROM reminders WHERE vehicle_id=? AND title=? AND completed=0", (vehicle_id, title)).fetchone()
        if not exists:
            conn.execute("INSERT INTO reminders(vehicle_id,reminder_date,reminder_type,title,notes) VALUES(?,?,?,?,?)", (vehicle_id, due_date, "Smart Reminder", title, notes))

    vehicles = conn.execute("SELECT * FROM vehicles WHERE status NOT IN ('Sold','BER')").fetchall()
    for v in vehicles:
        if v["rego_expiry"]:
            try:
                d = date.fromisoformat(v["rego_expiry"][:10])
                if today <= d <= today + timedelta(days=30):
                    add_once(v["id"], d.isoformat(), "Registration expires soon", f"{v['stock_no']} registration expires {d.isoformat()}.")
            except ValueError:
                pass
        if v["service_due_date"]:
            try:
                d = date.fromisoformat(v["service_due_date"][:10])
                if d <= today + timedelta(days=14):
                    add_once(v["id"], d.isoformat(), "Service due", f"{v['stock_no']} service is due {d.isoformat()}.")
            except ValueError:
                pass
        purchase_text = v["purchase_date"] or (str(v["created_at"] or "")[:10])
        try:
            bought = date.fromisoformat(purchase_text[:10])
            days = (today - bought).days
            if days >= 90:
                add_once(v["id"], today.isoformat(), "Stock aged 90+ days", f"{v['stock_no']} has been in stock {days} days. Review pricing/advertising.")
            elif days >= 60:
                add_once(v["id"], today.isoformat(), "Stock aged 60+ days", f"{v['stock_no']} has been in stock {days} days. Review next action.")
        except (ValueError, TypeError):
            pass
    drafts = conn.execute("""
        SELECT s.vehicle_id,s.sale_date,v.stock_no FROM sales s JOIN vehicles v ON v.id=s.vehicle_id
        WHERE COALESCE(s.invoice_status,'Draft')='Draft' AND COALESCE(s.sale_date,'')<>''
    """).fetchall()
    for row in drafts:
        try:
            d = date.fromisoformat(row["sale_date"][:10])
            if d <= today - timedelta(days=7):
                add_once(row["vehicle_id"], today.isoformat(), "Invoice still draft", f"Review the sale invoice for {row['stock_no']}.")
        except ValueError:
            pass
    reserved_parts = conn.execute("""
        SELECT id,vehicle_id,part_number,part_name,reserved_until,reserved_for
        FROM parts WHERE status='Reserved' AND vehicle_id IS NOT NULL AND COALESCE(reserved_until,'')<>''
    """).fetchall()
    for part in reserved_parts:
        try:
            d = date.fromisoformat(part["reserved_until"][:10])
            if d <= today + timedelta(days=2):
                add_once(part["vehicle_id"], d.isoformat(), f"Reserved part due: {part['part_number'] or part['part_name']}", f"Reservation for {part['reserved_for'] or 'customer'} expires {d.isoformat()}.")
        except ValueError:
            pass
    conn.commit()


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def owner_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        if session.get("role") != "owner":
            flash("Owner access is required.", "error")
            return redirect(url_for("dashboard"))
        return fn(*args, **kwargs)
    return wrapper

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def save_upload(file):
    if not file or not file.filename:
        return None
    if not allowed_file(file.filename):
        raise ValueError("Unsupported file type.")
    name = f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{secure_filename(file.filename)}"
    file.save(UPLOAD_DIR / name)
    return name

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"].strip().lower()
        password = request.form["password"]
        conn = db()
        user = conn.execute("SELECT * FROM users WHERE username=? AND COALESCE(is_active,1)=1", (username,)).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            conn.execute("UPDATE users SET last_login=CURRENT_TIMESTAMP WHERE id=?", (user["id"],))
            conn.commit()
            session.clear()
            session["user_id"] = user["id"]
            session["display_name"] = user["display_name"]
            session["role"] = user["role"]
            conn.close()
            return redirect(url_for("dashboard"))
        conn.close()
        flash("Incorrect username or password.", "error")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
@login_required
def dashboard():
    conn = db()
    ensure_smart_reminders(conn)

    metrics = conn.execute("""
        SELECT
          COUNT(*) AS total_vehicles,
          SUM(CASE WHEN status='Sold' THEN 1 ELSE 0 END) AS sold,
          SUM(CASE WHEN status='BER' THEN 1 ELSE 0 END) AS ber,
          SUM(CASE WHEN status NOT IN ('Sold','BER') THEN 1 ELSE 0 END) AS active_stock,
          SUM(CASE WHEN status NOT IN ('Sold','BER') THEN purchase_price_inc_gst ELSE 0 END) AS stock_value,
          SUM(purchase_price_inc_gst) AS purchase_total
        FROM vehicles
    """).fetchone()

    expense_total = conn.execute("SELECT COALESCE(SUM(cost_inc_gst),0) AS v FROM expenses").fetchone()["v"]
    job_total = conn.execute("SELECT COALESCE(SUM(CASE WHEN actual_cost_inc_gst>0 THEN actual_cost_inc_gst ELSE estimated_cost END),0) AS v FROM job_cards").fetchone()["v"]
    service_total = conn.execute("SELECT COALESCE(SUM(cost_inc_gst),0) AS v FROM service_entries").fetchone()["v"]
    parts_total = conn.execute("SELECT COALESCE(SUM(quantity_used*unit_cost_inc_gst),0) AS v FROM part_usage").fetchone()["v"]
    sales_total = conn.execute("SELECT COALESCE(SUM(sale_price_inc_gst),0) AS v FROM sales").fetchone()["v"]

    gst_paid = conn.execute("SELECT COALESCE(SUM(purchase_gst),0) AS v FROM vehicles").fetchone()["v"]
    gst_paid += conn.execute("SELECT COALESCE(SUM(gst_amount),0) AS v FROM expenses").fetchone()["v"]
    gst_paid += conn.execute("SELECT COALESCE(SUM(gst_amount),0) AS v FROM job_cards").fetchone()["v"]
    gst_paid += conn.execute("SELECT COALESCE(SUM(gst_amount),0) AS v FROM service_entries").fetchone()["v"]
    gst_collected = conn.execute("SELECT COALESCE(SUM(sale_gst),0) AS v FROM sales").fetchone()["v"]

    barry_total = conn.execute("SELECT COALESCE(SUM(barry_contribution),0) AS v FROM vehicles").fetchone()["v"]
    barry_total += conn.execute("SELECT COALESCE(SUM(CASE WHEN paid_by='Barry' THEN cost_inc_gst WHEN paid_by='Shared' THEN cost_inc_gst/2 ELSE 0 END),0) AS v FROM expenses").fetchone()["v"]
    barry_total += conn.execute(
        "SELECT COALESCE(SUM(CASE "
        "WHEN paid_by='Barry' THEN CASE WHEN actual_cost_inc_gst>0 THEN actual_cost_inc_gst ELSE estimated_cost END "
        "WHEN paid_by='Shared' THEN (CASE WHEN actual_cost_inc_gst>0 THEN actual_cost_inc_gst ELSE estimated_cost END)/2 "
        "ELSE 0 END),0) AS v FROM job_cards"
    ).fetchone()["v"]
    barry_total += conn.execute("SELECT COALESCE(SUM(CASE WHEN paid_by='Barry' THEN cost_inc_gst WHEN paid_by='Shared' THEN cost_inc_gst/2 ELSE 0 END),0) AS v FROM service_entries").fetchone()["v"]
    barry_total += conn.execute("SELECT COALESCE(SUM(CASE WHEN paid_by='Barry' THEN quantity_used*unit_cost_inc_gst WHEN paid_by='Shared' THEN quantity_used*unit_cost_inc_gst/2 ELSE 0 END),0) AS v FROM part_usage").fetchone()["v"]

    matt_total = conn.execute("SELECT COALESCE(SUM(matt_contribution),0) AS v FROM vehicles").fetchone()["v"]
    matt_total += conn.execute("SELECT COALESCE(SUM(CASE WHEN paid_by='Matt' THEN cost_inc_gst WHEN paid_by='Shared' THEN cost_inc_gst/2 ELSE 0 END),0) AS v FROM expenses").fetchone()["v"]
    matt_total += conn.execute(
        "SELECT COALESCE(SUM(CASE "
        "WHEN paid_by='Matt' THEN CASE WHEN actual_cost_inc_gst>0 THEN actual_cost_inc_gst ELSE estimated_cost END "
        "WHEN paid_by='Shared' THEN (CASE WHEN actual_cost_inc_gst>0 THEN actual_cost_inc_gst ELSE estimated_cost END)/2 "
        "ELSE 0 END),0) AS v FROM job_cards"
    ).fetchone()["v"]
    matt_total += conn.execute("SELECT COALESCE(SUM(CASE WHEN paid_by='Matt' THEN cost_inc_gst WHEN paid_by='Shared' THEN cost_inc_gst/2 ELSE 0 END),0) AS v FROM service_entries").fetchone()["v"]
    matt_total += conn.execute("SELECT COALESCE(SUM(CASE WHEN paid_by='Matt' THEN quantity_used*unit_cost_inc_gst WHEN paid_by='Shared' THEN quantity_used*unit_cost_inc_gst/2 ELSE 0 END),0) AS v FROM part_usage").fetchone()["v"]

    status_counts = conn.execute("""
        SELECT status, COUNT(*) AS total
        FROM vehicles
        GROUP BY status
        ORDER BY status
    """).fetchall()

    today = date.today()
    rego_30 = (today + timedelta(days=30)).isoformat()
    rego_60 = (today + timedelta(days=60)).isoformat()
    rego_90 = (today + timedelta(days=90)).isoformat()

    rego_alerts = conn.execute("""
        SELECT id,stock_no,make,model,registration,rego_expiry
        FROM vehicles
        WHERE rego_expiry IS NOT NULL AND rego_expiry!=''
          AND rego_expiry<=?
          AND status NOT IN ('Sold','BER')
        ORDER BY rego_expiry
        LIMIT 8
    """, (rego_90,)).fetchall()

    service_alerts = conn.execute("""
        SELECT id,stock_no,make,model,service_due_date
        FROM vehicles
        WHERE service_due_date IS NOT NULL AND service_due_date!=''
          AND service_due_date<=?
          AND status NOT IN ('Sold','BER')
        ORDER BY service_due_date
        LIMIT 8
    """, (rego_90,)).fetchall()

    recent_vehicles = conn.execute("""
        SELECT v.*,
          COALESCE((SELECT SUM(cost_inc_gst) FROM expenses e WHERE e.vehicle_id=v.id),0) AS expenses,
          COALESCE((SELECT SUM(CASE WHEN actual_cost_inc_gst>0 THEN actual_cost_inc_gst ELSE estimated_cost END)
                    FROM job_cards j WHERE j.vehicle_id=v.id),0) AS jobs,
          COALESCE((SELECT sale_price_inc_gst FROM sales s WHERE s.vehicle_id=v.id),0) AS sale_price
        FROM vehicles v
        ORDER BY v.id DESC
        LIMIT 8
    """).fetchall()

    monthly_rows = conn.execute("""
        SELECT substr(sale_date,1,7) AS month,
               COUNT(*) AS vehicles_sold,
               COALESCE(SUM(sale_price_inc_gst),0) AS sales_total
        FROM sales
        WHERE sale_date IS NOT NULL AND sale_date!=''
        GROUP BY substr(sale_date,1,7)
        ORDER BY month DESC
        LIMIT 12
    """).fetchall()
    monthly_rows = list(reversed(monthly_rows))

    next_stock = next_stock_number(conn)

    today_text = today.isoformat()
    open_tasks = conn.execute("""
        SELECT t.*,v.stock_no,v.make,v.model
        FROM tasks t
        LEFT JOIN vehicles v ON v.id=t.vehicle_id
        WHERE t.status!='Completed'
        ORDER BY CASE t.priority WHEN 'Urgent' THEN 1 WHEN 'High' THEN 2 ELSE 3 END,
                 COALESCE(t.due_date,'9999-12-31'),t.id
        LIMIT 8
    """).fetchall()

    attention_vehicles = conn.execute("""
        SELECT v.*,
          CAST(julianday(?) - julianday(COALESCE(v.purchase_date,substr(v.created_at,1,10))) AS INTEGER) AS days_in_stock,
          COALESCE((SELECT COUNT(*) FROM reminders r WHERE r.vehicle_id=v.id AND r.completed=0),0) AS open_reminders,
          COALESCE((SELECT COUNT(*) FROM tasks t WHERE t.vehicle_id=v.id AND t.status!='Completed'),0) AS open_tasks
        FROM vehicles v
        WHERE v.status NOT IN ('Sold','BER')
        ORDER BY open_reminders DESC,open_tasks DESC,days_in_stock DESC
        LIMIT 8
    """, (today_text,)).fetchall()

    advertised_count = conn.execute(
        "SELECT COUNT(*) AS c FROM vehicles WHERE status='Advertised'"
    ).fetchone()["c"]

    sold_this_month = conn.execute("""
        SELECT COUNT(*) AS c FROM sales
        WHERE substr(sale_date,1,7)=substr(?,1,7)
    """, (today_text,)).fetchone()["c"]

    parts_metrics = conn.execute("""
        SELECT
          COUNT(*) AS part_lines,
          COALESCE(SUM(quantity_on_hand),0) AS parts_on_hand,
          COALESCE(SUM(quantity_on_hand * unit_cost_inc_gst),0) AS parts_value,
          SUM(CASE WHEN quantity_on_hand <= reorder_level THEN 1 ELSE 0 END) AS low_stock
        FROM parts
    """).fetchone()

    missing_photo_count = conn.execute("""
        SELECT COUNT(*) AS c
        FROM vehicles
        WHERE status NOT IN ('Sold','BER')
          AND COALESCE(photo_filename,'')=''
          AND NOT EXISTS (SELECT 1 FROM vehicle_photos p WHERE p.vehicle_id=vehicles.id)
    """).fetchone()["c"]

    workshop_metrics = conn.execute("""
        SELECT
          COUNT(*) AS total_jobs,
          SUM(CASE WHEN status='Open' THEN 1 ELSE 0 END) AS open_jobs,
          SUM(CASE WHEN status='In Progress' THEN 1 ELSE 0 END) AS in_progress,
          SUM(CASE WHEN status='Waiting Parts' THEN 1 ELSE 0 END) AS waiting_parts,
          SUM(CASE WHEN status='Completed' THEN 1 ELSE 0 END) AS completed_jobs
        FROM job_cards
    """).fetchone()

    priority_tasks = conn.execute("""
        SELECT t.*,v.stock_no,v.make,v.model
        FROM tasks t
        LEFT JOIN vehicles v ON v.id=t.vehicle_id
        WHERE t.status!='Completed'
          AND (t.priority IN ('Urgent','High') OR COALESCE(t.due_date,'9999-12-31')<=?)
        ORDER BY CASE t.priority WHEN 'Urgent' THEN 1 WHEN 'High' THEN 2 ELSE 3 END,
                 COALESCE(t.due_date,'9999-12-31'),t.id
        LIMIT 6
    """, (today_text,)).fetchall()

    ready_to_advertise = conn.execute("""
        SELECT id,stock_no,year,make,model,status,asking_price
        FROM vehicles
        WHERE status='Ready for Sale'
        ORDER BY COALESCE(ready_for_sale_date,purchase_date,substr(created_at,1,10)),id
        LIMIT 6
    """).fetchall()

    parts_sales_month = conn.execute("""
        SELECT COALESCE(SUM(sale_price),0) AS revenue, COUNT(*) AS sales_count
        FROM part_sales
        WHERE substr(sale_date,1,7)=substr(?,1,7)
    """, (today_text,)).fetchone()

    parts_retail_value = conn.execute("""
        SELECT COALESCE(SUM(quantity_on_hand * selling_price),0) AS value
        FROM parts
    """).fetchone()["value"]

    asset_summary = conn.execute("""
        SELECT COALESCE(asset_type,'Car') AS asset_type,
               COUNT(*) AS total,
               SUM(CASE WHEN status NOT IN ('Sold','BER') THEN 1 ELSE 0 END) AS active,
               COALESCE(SUM(CASE WHEN status NOT IN ('Sold','BER') THEN purchase_price_inc_gst ELSE 0 END),0) AS stock_value
        FROM vehicles
        GROUP BY COALESCE(asset_type,'Car')
        ORDER BY asset_type
    """).fetchall()

    asset_totals = {row["asset_type"]: dict(row) for row in asset_summary}

    equipment_metrics = conn.execute("""
        SELECT
          COUNT(*) AS total_equipment,
          COALESCE(SUM(current_value),0) AS current_value,
          SUM(CASE WHEN status='Checked Out' THEN 1 ELSE 0 END) AS checked_out,
          SUM(CASE WHEN status IN ('Needs Service','Under Repair') OR condition IN ('Needs Repair','Unserviceable') THEN 1 ELSE 0 END) AS needs_attention,
          SUM(CASE WHEN COALESCE(photo_filename,'')='' THEN 1 ELSE 0 END) AS missing_photos
        FROM equipment
    """).fetchone()

    equipment_due = conn.execute("""
        SELECT id,equipment_no,name,next_service_date,calibration_due,test_tag_due
        FROM equipment
        WHERE (next_service_date IS NOT NULL AND next_service_date!='' AND next_service_date<=?)
           OR (calibration_due IS NOT NULL AND calibration_due!='' AND calibration_due<=?)
           OR (test_tag_due IS NOT NULL AND test_tag_due!='' AND test_tag_due<=?)
        ORDER BY COALESCE(next_service_date,calibration_due,test_tag_due),equipment_no
        LIMIT 5
    """, (rego_90,rego_90,rego_90)).fetchall()

    # Version 22 - Business Intelligence KPIs
    business_kpis = {}
    business_kpis["customers"] = conn.execute("""
        SELECT COUNT(*) AS c FROM (
            SELECT LOWER(TRIM(COALESCE(buyer_email,buyer_phone,buyer_name,''))) AS customer_key FROM sales WHERE COALESCE(buyer_name,'')<>''
            UNION
            SELECT LOWER(TRIM(COALESCE(customer_email,customer_phone,customer_name,''))) FROM part_sales WHERE COALESCE(customer_name,'')<>''
        ) WHERE customer_key<>''
    """).fetchone()["c"]
    business_kpis["smart_reminders"] = conn.execute("SELECT COUNT(*) AS c FROM reminders WHERE completed=0 AND reminder_type='Smart Reminder'").fetchone()["c"]
    business_kpis["donor_vehicles"] = conn.execute("SELECT COUNT(*) AS c FROM vehicles WHERE COALESCE(vehicle_purpose,'Retail Sale')='Parts Vehicle' OR status='BER'").fetchone()["c"]
    business_kpis["inventory_retail"] = parts_retail_value
    business_kpis["parts_month"] = float(parts_sales_month["revenue"] or 0)
    business_kpis["vehicle_month"] = float(conn.execute("SELECT COALESCE(SUM(sale_price_inc_gst),0) AS v FROM sales WHERE substr(sale_date,1,7)=substr(?,1,7)", (today_text,)).fetchone()["v"] or 0)
    business_kpis["avg_stock_days"] = float(conn.execute("""
        SELECT COALESCE(AVG(julianday(?) - julianday(COALESCE(purchase_date,substr(created_at,1,10)))),0) AS v
        FROM vehicles WHERE status NOT IN ('Sold','BER')
    """, (today_text,)).fetchone()["v"] or 0)
    business_kpis["advertised"] = advertised_count

    email_metrics = conn.execute("""
        SELECT COUNT(*) AS total,
               SUM(CASE WHEN status='Unread' THEN 1 ELSE 0 END) AS unread,
               SUM(CASE WHEN status='Follow Up' THEN 1 ELSE 0 END) AS follow_up
        FROM email_messages
    """).fetchone()
    recent_emails = conn.execute("""
        SELECT em.*, v.stock_no, c.name AS contact_name
        FROM email_messages em
        LEFT JOIN vehicles v ON v.id=em.vehicle_id
        LEFT JOIN contacts c ON c.id=em.contact_id
        ORDER BY COALESCE(em.message_date, substr(em.created_at,1,10)) DESC, em.id DESC
        LIMIT 5
    """).fetchall()

    conn.close()

    net_profit = sales_total - (metrics["purchase_total"] or 0) - expense_total - job_total - service_total - parts_total

    status_labels = [row["status"] for row in status_counts]
    status_values = [row["total"] for row in status_counts]
    month_labels = [row["month"] for row in monthly_rows]
    month_values = [row["sales_total"] for row in monthly_rows]

    return render_template(
        "dashboard.html",
        metrics=metrics,
        expense_total=expense_total,
        job_total=job_total,
        service_total=service_total,
        parts_total=parts_total,
        sales_total=sales_total,
        gst_paid=gst_paid,
        gst_collected=gst_collected,
        net_profit=net_profit,
        barry_total=barry_total,
        matt_total=matt_total,
        rego_alerts=rego_alerts,
        service_alerts=service_alerts,
        recent_vehicles=recent_vehicles,
        next_stock=next_stock,
        status_labels=status_labels,
        status_values=status_values,
        month_labels=month_labels,
        month_values=month_values,
        today=today.isoformat(),
        open_tasks=open_tasks,
        attention_vehicles=attention_vehicles,
        advertised_count=advertised_count,
        sold_this_month=sold_this_month,
        parts_metrics=parts_metrics,
        missing_photo_count=missing_photo_count,
        workshop_metrics=workshop_metrics,
        priority_tasks=priority_tasks,
        ready_to_advertise=ready_to_advertise,
        parts_sales_month=parts_sales_month,
        parts_retail_value=parts_retail_value,
        asset_summary=asset_summary,
        asset_totals=asset_totals,
        equipment_metrics=equipment_metrics,
        equipment_due=equipment_due,
        business_kpis=business_kpis,
        email_metrics=email_metrics,
        recent_emails=recent_emails,
    )
@app.route("/partner-settlement")
@login_required
def partner_settlement():
    conn = db()

    vehicles = conn.execute("""
        SELECT
            id,
            stock_no,
            year,
            make,
            model,
            status,
            purchase_price_inc_gst,
            barry_contribution,
            matt_contribution
        FROM vehicles
        ORDER BY id DESC
    """).fetchall()

    barry_total = sum(float(v["barry_contribution"] or 0) for v in vehicles)
    matt_total = sum(float(v["matt_contribution"] or 0) for v in vehicles)

    difference = barry_total - matt_total
    settlement_amount = abs(difference) / 2

    if difference > 0:
        settlement_message = f"Matt owes Barry ${settlement_amount:,.2f}"
    elif difference < 0:
        settlement_message = f"Barry owes Matt ${settlement_amount:,.2f}"
    else:
        settlement_message = "Barry and Matt are fully balanced"

    conn.close()

    return render_template(
        "partner_settlement.html",
        vehicles=vehicles,
        barry_total=barry_total,
        matt_total=matt_total,
        difference=difference,
        settlement_amount=settlement_amount,
        settlement_message=settlement_message,
    )
@app.route("/vehicles")
@login_required
def vehicle_list():
    q = request.args.get("q", "").strip()
    asset_type = request.args.get("asset_type", "").strip()
    conn = db()
    base_query = """
        SELECT v.*,
               COALESCE(
                 v.photo_filename,
                 (SELECT fp.filename FROM vehicle_photos fp WHERE fp.id=v.featured_photo_id LIMIT 1),
                 (SELECT p.filename FROM vehicle_photos p WHERE p.vehicle_id=v.id ORDER BY p.id DESC LIMIT 1)
               ) AS thumbnail
        FROM vehicles v
    """
    conditions = []
    params = []
    if q:
        conditions.append("(v.stock_no LIKE ? OR v.make LIKE ? OR v.model LIKE ? OR v.vin LIKE ? OR v.registration LIKE ? OR v.hin LIKE ?)")
        params.extend([f"%{q}%"] * 6)
    if asset_type:
        conditions.append("COALESCE(v.asset_type,'Car')=?")
        params.append(asset_type)
    if conditions:
        base_query += " WHERE " + " AND ".join(conditions)
    rows = conn.execute(base_query + " ORDER BY v.id DESC", tuple(params)).fetchall()
    type_counts = conn.execute("""
        SELECT COALESCE(asset_type,'Car') AS asset_type, COUNT(*) AS total
        FROM vehicles GROUP BY COALESCE(asset_type,'Car')
    """).fetchall()
    conn.close()
    return render_template("vehicles.html", vehicles=rows, q=q, asset_type=asset_type, type_counts=type_counts)

@app.route("/vehicles/new", methods=["GET", "POST"])
@login_required
def vehicle_new():
    conn = db()
    suggested_stock = next_stock_number(conn)
    conn.close()
    if request.method == "POST":
        try:
            photo = save_upload(request.files.get("photo"))
            price = float(request.form.get("purchase_price_inc_gst") or 0)
            gst = round(price / 11, 2)
            conn = db()
            stock_no = request.form.get("stock_no", "").strip() or next_stock_number(conn)
            duplicate_rows = duplicate_vehicle_matches(conn, request.form.get("vin"), request.form.get("registration"), stock_no)
            if duplicate_rows:
                hit = duplicate_rows[0]
                raise ValueError(f"Possible duplicate stock: {hit['stock_no']} {hit['year'] or ''} {hit['make']} {hit['model']} (VIN/registration/stock match).")
            fields = [
                "stock_no","status","purchase_date","make","model","variant","year","vin","registration",
                "odometer_km","colour","purchase_price_inc_gst","purchase_gst","barry_contribution",
                "matt_contribution","sale_ownership","rego_expiry","photo_filename","notes","ppsr_number","roadworthy_status",
                "service_due_date","service_history","asset_type","length_m","width_m","tare_Weight_kg","atm_kg","gtm_kg",
                "berths","axles","caravan_features","trailer_features","boat_type","hull_material","engine_make","engine_model",
                "engine_hours","horsepower","fuel_type","hin","trailer_included","trailer_registration",
                "capacity_people","boat_features","vehicle_purpose","dismantling_status"
            ]
            values = [
                stock_no, request.form.get("status") or "In Stock", request.form.get("purchase_date"),
                (request.form.get("make") or "").strip(), (request.form.get("model") or "").strip(),
                request.form.get("variant"), request.form.get("year") or None, request.form.get("vin") or None,
                request.form.get("registration"), request.form.get("odometer_km") or None, request.form.get("colour"),
                price, gst, float(request.form.get("barry_contribution") or 0),
                float(request.form.get("matt_contribution") or 0),request.form.get("sale_ownership") or "BAM Joint", request.form.get("rego_expiry"), photo,
                request.form.get("notes"), request.form.get("ppsr_number"),
                request.form.get("roadworthy_status") or "Not Checked", request.form.get("service_due_date"),
                request.form.get("service_history"), request.form.get("asset_type") or "Car",
                request.form.get("length_m") or None, request.form.get("width_m") or None, request.form.get("tare_weight_kg") or None,
                request.form.get("atm_kg") or None, request.form.get("gtm_kg") or None,
                request.form.get("berths") or None, request.form.get("axles") or None,
                request.form.get("caravan_features"), request.form.get("trailer_features"), request.form.get("boat_type"),
                request.form.get("hull_material"), request.form.get("engine_make"),
                request.form.get("engine_model"), request.form.get("engine_hours") or None,
                request.form.get("horsepower") or None, request.form.get("fuel_type"),
                request.form.get("hin"), 1 if request.form.get("trailer_included") else 0,
                request.form.get("trailer_registration"), request.form.get("capacity_people") or None,
                request.form.get("boat_features"),
                request.form.get("vehicle_purpose") or ("Parts Vehicle" if (request.form.get("status") == "BER") else "Retail Sale"),
                request.form.get("dismantling_status") or "Not Started"
            ]
            if not values[3] or not values[4]:
                raise ValueError("Make and model are required.")
            placeholders = ",".join(["?"] * len(fields))
            conn.execute(f"INSERT INTO vehicles({','.join(fields)}) VALUES({placeholders})", values)
            conn.commit()
            conn.close()
            flash(f"{request.form.get('asset_type') or 'Asset'} added.", "success")
            return redirect(url_for("vehicle_list"))
        except (sqlite3.IntegrityError, ValueError) as exc:
            flash(str(exc), "error")
    return render_template("vehicle_form.html", suggested_stock=suggested_stock, vehicle_makes=VEHICLE_MAKES, vehicle_model_catalog=VEHICLE_MODEL_CATALOG, fuel_types=FUEL_TYPES, transmission_types=TRANSMISSION_TYPES, drive_types=DRIVE_TYPES)

@app.route("/vehicles/<int:vehicle_id>/edit", methods=["GET", "POST"])
@login_required
def vehicle_edit(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close()
        return "Vehicle not found", 404

    if request.method == "POST":
        try:
            old_values = dict(vehicle)
            price = float(request.form.get("purchase_price_inc_gst") or 0)
            gst = round(price / 11, 2)
            new_photo = save_upload(request.files.get("photo"))
            photo_filename = new_photo or vehicle["photo_filename"]

            values = {
                "stock_no": (request.form.get("stock_no") or vehicle["stock_no"]).strip(),
                "status": request.form.get("status") or "In Stock",
                "purchase_date": request.form.get("purchase_date") or None,
                "make": (request.form.get("make") or "").strip(),
                "model": (request.form.get("model") or "").strip(),
                "variant": request.form.get("variant") or None,
                "year": request.form.get("year") or None,
                "vin": (request.form.get("vin") or "").strip() or None,
                "registration": (request.form.get("registration") or "").strip() or None,
                "odometer_km": request.form.get("odometer_km") or None,
                "colour": request.form.get("colour") or None,
                "purchase_price_inc_gst": price,
                "purchase_gst": gst,
                "asking_price": float(request.form.get("asking_price") or 0),
                "minimum_sale_price": float(request.form.get("minimum_sale_price") or 0),
                "negotiated_price": float(request.form.get("negotiated_price") or 0),
                "barry_contribution": float(request.form.get("barry_contribution") or 0),
                "matt_contribution": float(request.form.get("matt_contribution") or 0),
                "sale_ownership": request.form.get("sale_ownership") or vehicle["sale_ownership"] or "BAM Joint",
                "rego_expiry": request.form.get("rego_expiry") or None,
                "photo_filename": photo_filename,
                "notes": request.form.get("notes") or None,
                "ppsr_number": request.form.get("ppsr_number") or None,
                "roadworthy_status": request.form.get("roadworthy_status") or "Not Checked",
                "service_due_date": request.form.get("service_due_date") or None,
                "service_history": request.form.get("service_history") or None,
                "asset_type": request.form.get("asset_type") or "Car",
                "length_m": request.form.get("length_m") or None,
                "width_m": request.form.get("width_m") or None,
                "tare_weight_kg": request.form.get("tare_weight_kg") or None,
                "atm_kg": request.form.get("atm_kg") or None,
                "gtm_kg": request.form.get("gtm_kg") or None,
                "berths": request.form.get("berths") or None,
                "axles": request.form.get("axles") or None,
                "caravan_features": request.form.get("caravan_features") or None,
                "trailer_features": request.form.get("trailer_features") or None,
                "boat_type": request.form.get("boat_type") or None,
                "hull_material": request.form.get("hull_material") or None,
                "engine_make": request.form.get("engine_make") or None,
                "engine_model": request.form.get("engine_model") or None,
                "engine_hours": request.form.get("engine_hours") or None,
                "horsepower": request.form.get("horsepower") or None,
                "fuel_type": request.form.get("fuel_type") or None,
                "hin": request.form.get("hin") or None,
                "trailer_included": 1 if request.form.get("trailer_included") else 0,
                "trailer_registration": request.form.get("trailer_registration") or None,
                "capacity_people": request.form.get("capacity_people") or None,
                "boat_features": request.form.get("boat_features") or None,
                "vehicle_purpose": request.form.get("vehicle_purpose") or ("Parts Vehicle" if (request.form.get("status") == "BER") else "Retail Sale"),
                "dismantling_status": request.form.get("dismantling_status") or "Not Started",
                "shell_sale_price": float(request.form.get("shell_sale_price") or 0),
                "shell_sale_date": request.form.get("shell_sale_date") or None,
            }
            if not values["make"] or not values["model"]:
                raise ValueError("Make and model are required.")
            duplicate_rows = duplicate_vehicle_matches(conn, values["vin"], values["registration"], values["stock_no"], exclude_id=vehicle_id)
            if duplicate_rows:
                hit = duplicate_rows[0]
                raise ValueError(f"Possible duplicate stock: {hit['stock_no']} {hit['year'] or ''} {hit['make']} {hit['model']} (VIN/registration/stock match).")

            conn.execute("""
                UPDATE vehicles SET
                    stock_no=?, status=?, purchase_date=?, make=?, model=?, variant=?, year=?,
                    vin=?, registration=?, odometer_km=?, colour=?, purchase_price_inc_gst=?,
                    purchase_gst=?, asking_price=?, minimum_sale_price=?, negotiated_price=?,
                    barry_contribution=?, matt_contribution=?, sale_ownership=?, rego_expiry=?,
                    photo_filename=?, notes=?, ppsr_number=?, roadworthy_status=?,
                    service_due_date=?, service_history=?, asset_type=?, length_m=?, tare_weight_kg=?,
                    atm_kg=?, gtm_kg=?, berths=?, axles=?, caravan_features=?, boat_type=?, hull_material=?,
                    engine_make=?, engine_model=?, engine_hours=?, horsepower=?, fuel_type=?, hin=?,
                    trailer_included=?, trailer_registration=?, capacity_people=?, boat_features=?,
                    vehicle_purpose=?, dismantling_status=?, shell_sale_price=?, shell_sale_date=?
                WHERE id=?
            """, (
                values["stock_no"], values["status"], values["purchase_date"],
                values["make"], values["model"], values["variant"], values["year"],
                values["vin"], values["registration"], values["odometer_km"],
                values["colour"], values["purchase_price_inc_gst"], values["purchase_gst"],
                values["asking_price"], values["minimum_sale_price"], values["negotiated_price"],
                values["barry_contribution"], values["matt_contribution"],
                values["sale_ownership"],
                values["rego_expiry"], values["photo_filename"], values["notes"],
                values["ppsr_number"], values["roadworthy_status"],
                values["service_due_date"], values["service_history"], values["asset_type"],
                values["length_m"], values["tare_weight_kg"], values["atm_kg"], values["gtm_kg"],
                values["berths"], values["axles"], values["caravan_features"], values["boat_type"],
                values["hull_material"], values["engine_make"], values["engine_model"],
                values["engine_hours"], values["horsepower"], values["fuel_type"], values["hin"],
                values["trailer_included"], values["trailer_registration"], values["capacity_people"],
                values["boat_features"], values["vehicle_purpose"], values["dismantling_status"],
                values["shell_sale_price"], values["shell_sale_date"], vehicle_id,
            ))
            if values["asking_price"] > 0 and values["minimum_sale_price"] > 0:
                conn.execute("""
                    UPDATE tasks
                    SET status='Completed'
                    WHERE vehicle_id=?
                      AND title='Confirm asking price and minimum sale price'
                      AND status!='Completed'
                """, (vehicle_id,))
            conn.commit()
            conn.close()

            changed = []
            watched = [
                "stock_no", "status", "purchase_date", "make", "model", "variant",
                "year", "vin", "registration", "odometer_km", "colour",
                "purchase_price_inc_gst", "asking_price", "minimum_sale_price", "negotiated_price",
                "barry_contribution", "matt_contribution", "sale_ownership",
                "rego_expiry", "ppsr_number", "roadworthy_status",
                "service_due_date", "service_history", "asset_type", "length_m", "tare_weight_kg",
                "atm_kg", "gtm_kg", "berths", "axles", "caravan_features", "boat_type",
                "hull_material", "engine_make", "engine_model", "engine_hours", "horsepower",
                "fuel_type", "hin", "trailer_included", "trailer_registration", "capacity_people",
                "boat_features", "vehicle_purpose", "dismantling_status", "shell_sale_price",
                "shell_sale_date", "notes",
            ]
            for field in watched:
                old = old_values.get(field)
                new = values.get(field)
                if str(old or "") != str(new or ""):
                    changed.append(f"{field}: {old or ''} -> {new or ''}")
            if new_photo:
                changed.append("main photo replaced")
            log_action(
                "Vehicle updated",
                "vehicle",
                vehicle_id,
                "; ".join(changed) if changed else "Vehicle saved with no field changes",
            )
            flash("Vehicle updated.", "success")
            if request.form.get("continue_editing"):
                return redirect(url_for("vehicle_edit", vehicle_id=vehicle_id))
            return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))
        except (sqlite3.IntegrityError, ValueError) as exc:
            conn.close()
            flash(str(exc), "error")
            vehicle = {**dict(vehicle), **request.form.to_dict()}
 
    else:
        conn.close()

    return render_template("vehicle_edit.html", vehicle=vehicle)

@app.errorhandler(413)
def upload_too_large(_error):
    flash("The selected files are too large. Upload fewer photos at a time or use smaller files.", "error")
    return redirect(request.referrer or url_for("vehicle_list"))

@app.route("/api/vin-decode")
@login_required
def api_vin_decode():
    specs, error = decode_vin(request.args.get("vin", ""))
    return jsonify({"specs": specs, "error": error})


@app.route("/api/vehicle-duplicate-check")
@login_required
def vehicle_duplicate_check():
    conn = db()
    rows = duplicate_vehicle_matches(conn, request.args.get("vin"), request.args.get("registration"), request.args.get("stock_no"), request.args.get("exclude_id", type=int))
    conn.close()
    return jsonify({"matches": [dict(r) for r in rows]})



@app.route("/vin-intelligence")
@login_required
def vin_intelligence_centre():
    q = (request.args.get("q") or "").strip()
    conn = db()
    params = []
    where = ""
    if q:
        like = f"%{q}%"
        where = "WHERE stock_no LIKE ? OR vin LIKE ? OR registration LIKE ? OR make LIKE ? OR model LIKE ?"
        params = [like] * 5
    vehicles = conn.execute(f"""SELECT id,stock_no,year,make,model,variant,vin,registration,status,decoded_at,
        manufacturer_name,engine_code,transmission_code,paint_code FROM vehicles {where} ORDER BY id DESC LIMIT 80""", params).fetchall()
    conn.close()
    return render_template("vin_intelligence_centre.html", vehicles=vehicles, q=q, vin_provider=VIN_DATA_PROVIDER_NAME, provider_name=WORKSHOP_PROVIDER_NAME)


@app.route("/vehicles/<int:vehicle_id>/intelligence")
@login_required
def dealer_intelligence(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close()
        return "Vehicle not found", 404
    market = market_price_suggestion(conn, vehicle)
    duplicates = duplicate_vehicle_matches(conn, vehicle["vin"], vehicle["registration"], vehicle["stock_no"], exclude_id=vehicle_id)
    comparable_sales = conn.execute("""
        SELECT s.sale_date,s.sale_price_inc_gst,v.stock_no,v.year,v.make,v.model
        FROM sales s JOIN vehicles v ON v.id=s.vehicle_id
        WHERE s.sale_price_inc_gst>0 AND LOWER(v.make)=LOWER(?)
        ORDER BY s.sale_date DESC LIMIT 8
    """, (vehicle["make"],)).fetchall()
    conn.close()
    return render_template("dealer_intelligence.html", vehicle=vehicle, market=market, duplicates=duplicates,
                           comparable_sales=comparable_sales, ai_enabled=bool(OPENAI_API_KEY))


@app.route("/vehicles/<int:vehicle_id>/decode-vin", methods=["POST"])
@login_required
def vehicle_decode_vin(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close()
        return "Vehicle not found", 404
    specs, error = decode_vin(vehicle["vin"])
    if specs:
        conn.execute("""
            UPDATE vehicles SET vin_decode_json=?,decoded_at=CURRENT_TIMESTAMP,manufacturer_name=?,body_class=?,drive_type=?,
                transmission_style=?,engine_cylinders=?,engine_displacement_l=?,fuel_type_primary=?,decoded_model_year=?,decoded_series=?,decoded_trim=?,
                vehicle_type_decoded=?,doors_decoded=?,engine_model_decoded=?,engine_power_kw_decoded=?,plant_country_decoded=?,
                brake_system_decoded=?,turbo_decoded=?,electrification_level_decoded=?,
                make=CASE WHEN COALESCE(make,'')='' THEN ? ELSE make END,
                model=CASE WHEN COALESCE(model,'')='' THEN ? ELSE model END,
                year=CASE WHEN year IS NULL OR year=0 THEN ? ELSE year END
            WHERE id=?
        """, (json.dumps(specs.get("raw", {})), specs.get("manufacturer_name"), specs.get("body_class"), specs.get("drive_type"),
              specs.get("transmission_style"), specs.get("engine_cylinders"), specs.get("engine_displacement_l"), specs.get("fuel_type_primary"),
              specs.get("year"), specs.get("series"), specs.get("trim"), specs.get("vehicle_type"), specs.get("doors"),
              specs.get("engine_model"), specs.get("engine_power_kw"), specs.get("plant_country"), specs.get("brake_system"),
              specs.get("turbo"), specs.get("electrification_level"), specs.get("make"), specs.get("model"), specs.get("year") or None, vehicle_id))
        conn.commit()
        log_action("VIN decoded", "vehicle", vehicle_id, f"VIN {vehicle['vin']}; provider specs saved")
        flash("VIN decoded and available specifications were saved." + (f" Note: {error}" if error else ""), "success")
    else:
        flash(error or "VIN could not be decoded.", "error")
    conn.close()
    target = request.form.get("return_to") or request.referrer or url_for("dealer_intelligence", vehicle_id=vehicle_id)
    return redirect(target)


@app.route("/vehicles/<int:vehicle_id>/vin-intelligence")
@login_required
def vin_intelligence(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close(); return "Vehicle not found", 404
    refs = conn.execute("SELECT * FROM workshop_references WHERE vehicle_id=? ORDER BY reference_type,title", (vehicle_id,)).fetchall()
    operations = conn.execute("SELECT * FROM workshop_operations WHERE vehicle_id=? ORDER BY system_name,operation_name LIMIT 30", (vehicle_id,)).fetchall()
    duplicates = duplicate_vehicle_matches(conn, vehicle["vin"], vehicle["registration"], vehicle["stock_no"], exclude_id=vehicle_id)
    conn.close()
    fields = [
        vehicle["vin"], vehicle["make"], vehicle["model"], vehicle["year"], vehicle["manufacturer_name"],
        vehicle["body_class"], vehicle["drive_type"], vehicle["transmission_style"], vehicle["engine_displacement_l"],
        vehicle["fuel_type_primary"], vehicle["engine_code"], vehicle["transmission_code"], vehicle["paint_code"]
    ]
    completeness = round(sum(1 for x in fields if str(x or '').strip()) / len(fields) * 100)
    manual_search_phrase = " ".join(str(x).strip() for x in [vehicle["year"], vehicle["make"], vehicle["model"], vehicle["variant"], vehicle["vin"], "workshop service repair manual"] if x)
    return render_template("vin_intelligence.html", vehicle=vehicle, references=refs, operations=operations,
                           duplicates=duplicates, completeness=completeness, provider_name=WORKSHOP_PROVIDER_NAME,
                           provider_url=_workshop_provider_vehicle_url(vehicle), vin_provider=VIN_DATA_PROVIDER_NAME,
                           manual_search_phrase=manual_search_phrase)


@app.route("/vehicles/<int:vehicle_id>/vin-intelligence/specifications", methods=["POST"])
@login_required
def vin_intelligence_specifications_save(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close(); return "Vehicle not found", 404
    def optional_int(name):
        raw = (request.form.get(name) or "").strip()
        if not raw: return None
        value = int(raw)
        if value < 0: raise ValueError(f"{name} cannot be negative.")
        return value
    try:
        conn.execute("""UPDATE vehicles SET paint_code=?,engine_code=?,transmission_code=?,option_codes=?,
            service_interval_km=?,service_interval_months=?,technical_data_source=?,technical_notes=? WHERE id=?""", (
            (request.form.get("paint_code") or "").strip() or None,
            (request.form.get("engine_code") or "").strip() or None,
            (request.form.get("transmission_code") or "").strip() or None,
            (request.form.get("option_codes") or "").strip() or None,
            optional_int("service_interval_km"), optional_int("service_interval_months"),
            (request.form.get("technical_data_source") or "").strip() or None,
            (request.form.get("technical_notes") or "").strip() or None, vehicle_id
        ))
        conn.commit(); conn.close()
        log_action("Verified workshop specifications updated", "vehicle", vehicle_id, vehicle["stock_no"])
        flash("Verified workshop specifications saved.", "success")
    except (ValueError, sqlite3.Error) as exc:
        conn.rollback(); conn.close(); flash(str(exc), "error")
    return redirect(url_for("vin_intelligence", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/market-suggestion", methods=["POST"])
@login_required
def vehicle_market_suggestion_save(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close(); return "Vehicle not found", 404
    market = market_price_suggestion(conn, vehicle)
    conn.execute("UPDATE vehicles SET market_price_low=?,market_price_mid=?,market_price_high=?,market_price_checked_at=CURRENT_TIMESTAMP WHERE id=?",
                 (market["low"], market["mid"], market["high"], vehicle_id))
    conn.commit(); conn.close()
    log_action("Market pricing suggestion refreshed", "vehicle", vehicle_id, json.dumps(market))
    flash("Pricing suggestion refreshed from BAM's recorded sales and vehicle cost data.", "success")
    return redirect(url_for("dealer_intelligence", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/generate-ai-ad", methods=["POST"])
@login_required
def vehicle_generate_ai_ad(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close(); return "Vehicle not found", 404
    text, source = generate_vehicle_ad_text(vehicle)
    title = f"{vehicle['year'] or ''} {vehicle['make']} {vehicle['model']} {vehicle['variant'] or ''}".strip()
    conn.execute("UPDATE vehicles SET advertisement_title=?,advertisement_description=? WHERE id=?", (title, text, vehicle_id))
    conn.commit(); conn.close()
    log_action("Advertisement generated", "vehicle", vehicle_id, source)
    flash(f"Advertisement generated. {source}", "success")
    return redirect(url_for("advertisement_pro", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>")
@login_required
def vehicle_detail(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close()
        return "Asset not found", 404

    expenses = conn.execute(
        "SELECT * FROM expenses WHERE vehicle_id=? ORDER BY expense_date DESC,id DESC",
        (vehicle_id,),
    ).fetchall()
    job_cards = conn.execute(
        "SELECT * FROM job_cards WHERE vehicle_id=? ORDER BY job_date DESC,id DESC",
        (vehicle_id,),
    ).fetchall()
    photos = conn.execute(
        """SELECT * FROM vehicle_photos
           WHERE vehicle_id=?
           ORDER BY CASE WHEN id=(SELECT featured_photo_id FROM vehicles WHERE id=?) THEN 0 ELSE 1 END,
                    id DESC""",
        (vehicle_id, vehicle_id),
    ).fetchall()
    documents = conn.execute(
        "SELECT * FROM vehicle_documents WHERE vehicle_id=? ORDER BY uploaded_at DESC,id DESC",
        (vehicle_id,),
    ).fetchall()
    reminders = conn.execute(
        "SELECT * FROM reminders WHERE vehicle_id=? ORDER BY completed,reminder_date,id DESC",
        (vehicle_id,),
    ).fetchall()
    services = conn.execute(
        "SELECT * FROM service_entries WHERE vehicle_id=? ORDER BY service_date DESC,id DESC",
        (vehicle_id,),
    ).fetchall()
    parts_used = conn.execute(
        """SELECT u.*,p.part_number,p.part_name
           FROM part_usage u JOIN parts p ON p.id=u.part_id
           WHERE u.vehicle_id=? ORDER BY u.usage_date DESC,u.id DESC""",
        (vehicle_id,),
    ).fetchall()
    available_parts = conn.execute(
        "SELECT * FROM parts WHERE quantity_on_hand>0 ORDER BY part_name",
    ).fetchall()
    sale = conn.execute("SELECT * FROM sales WHERE vehicle_id=?", (vehicle_id,)).fetchone()

    expense_total = sum(float(row["cost_inc_gst"] or 0) for row in expenses)
    job_total = sum(
        float(row["actual_cost_inc_gst"] or 0)
        if float(row["actual_cost_inc_gst"] or 0) > 0
        else float(row["estimated_cost"] or 0)
        for row in job_cards
    )
    service_total = sum(float(row["cost_inc_gst"] or 0) for row in services)
    parts_total = sum(float(row["quantity_used"] or 0) * float(row["unit_cost_inc_gst"] or 0) for row in parts_used)
    sale_price = float(sale["sale_price_inc_gst"] or 0) if sale else 0
    selling_costs = float((sale["advertising_cost"] or 0) + (sale["transfer_cost"] or 0)) if sale else 0
    total_invested = float(vehicle["purchase_price_inc_gst"] or 0) + expense_total + job_total + service_total + parts_total + selling_costs

    def split_partner_costs(rows, amount_func):
        barry = 0.0
        matt = 0.0

        for row in rows:
            amount = float(amount_func(row) or 0)
            paid_by = str(row["paid_by"] or "Shared").strip().lower()

            if paid_by == "barry":
                barry += amount
            elif paid_by == "matt":
                matt += amount
            else:
                barry += amount / 2
                matt += amount / 2

        return barry, matt

    barry_invested = float(vehicle["barry_contribution"] or 0)
    matt_invested = float(vehicle["matt_contribution"] or 0)

    purchase_price = float(vehicle["purchase_price_inc_gst"] or 0)
    purchase_balance = max(
        purchase_price - barry_invested - matt_invested,
        0
    )
    sale_ownership = vehicle["sale_ownership"] or "BAM Joint"

    if sale_ownership == "Barry Personal":
        barry_invested += purchase_balance
    elif sale_ownership == "Matt Personal":
        matt_invested += purchase_balance
    else:
        barry_invested += purchase_balance / 2
        matt_invested += purchase_balance / 2

    barry_expenses, matt_expenses = split_partner_costs(
        expenses,
        lambda row: row["cost_inc_gst"]
    )

    barry_jobs, matt_jobs = split_partner_costs(
        job_cards,
        lambda row: (
            row["actual_cost_inc_gst"]
            if float(row["actual_cost_inc_gst"] or 0) > 0
            else row["estimated_cost"]
        )
    )

    barry_services, matt_services = split_partner_costs(
        services,
        lambda row: row["cost_inc_gst"]
    )

    barry_parts, matt_parts = split_partner_costs(
        parts_used,
        lambda row: float(row["quantity_used"] or 0)
        * float(row["unit_cost_inc_gst"] or 0)
    )

    barry_invested += (
        barry_expenses
        + barry_jobs
        + barry_services
        + barry_parts
        + selling_costs / 2
    )

    matt_invested += (
        matt_expenses
        + matt_jobs
        + matt_services
        + matt_parts
        + selling_costs / 2
    )

    vehicle_profit = sale_price - total_invested

    if sale and sale_price > 0:
        if sale_ownership == "Barry Personal":
            barry_receives = sale_price
            matt_receives = 0.0
        elif sale_ownership == "Matt Personal":
           barry_receives = 0.0
           matt_receives = sale_price
        else:
           barry_receives = barry_invested + vehicle_profit / 2
           matt_receives = matt_invested + vehicle_profit / 2
    else:
        barry_receives = 0.0
        matt_receives = 0.0
    # Version 21 - every dismantled-part sale feeds back to the donor vehicle.
    donor_parts = conn.execute(
        "SELECT * FROM parts WHERE vehicle_id=? ORDER BY id DESC", (vehicle_id,)
    ).fetchall()
    donor_metrics = conn.execute("""
        SELECT COUNT(*) AS part_lines,
               COALESCE(SUM(quantity_on_hand),0) AS units_on_hand,
               COALESCE(SUM(quantity_on_hand * selling_price),0) AS remaining_retail
        FROM parts WHERE vehicle_id=?
    """, (vehicle_id,)).fetchone()
    donor_sales = conn.execute("""
        SELECT COALESCE(SUM(ps.sale_price),0) AS revenue,
               COALESCE(SUM(ps.quantity),0) AS units_sold,
               COUNT(ps.id) AS sale_count
        FROM part_sales ps JOIN parts p ON p.id=ps.part_id
        WHERE p.vehicle_id=?
    """, (vehicle_id,)).fetchone()
    donor_revenue = float(donor_sales["revenue"] or 0)
    shell_revenue = float(vehicle["shell_sale_price"] or 0)
    total_revenue = sale_price + donor_revenue + shell_revenue
    profit = total_revenue - total_invested
    donor_profit = donor_revenue + shell_revenue - total_invested
    recovery_pct = (total_revenue / total_invested * 100) if total_invested > 0 else (100.0 if total_revenue > 0 else 0.0)
    remaining_to_break_even = max(0.0, total_invested - total_revenue)
    break_even_status = "Recovered / Profitable" if total_revenue >= total_invested and total_invested > 0 else "Recovering Investment"
    conn.close()

    vehicle_page = render_template(
        "vehicle_detail.html",
        vehicle=vehicle,
        expenses=expenses,
        job_cards=job_cards,
        photos=photos,
        documents=documents,
        reminders=reminders,
        services=services,
        parts_used=parts_used,
        available_parts=available_parts,
        sale=sale,
        expense_total=expense_total,
        job_total=job_total,
        service_total=service_total,
        parts_total=parts_total,
        total_invested=total_invested,
        profit=profit,
        barry_invested=barry_invested,
        matt_invested=matt_invested,
        barry_receives=barry_receives,
        matt_receives=matt_receives,
        vehicle_profit=vehicle_profit,
        donor_parts=donor_parts,
        donor_metrics=donor_metrics,
        donor_revenue=donor_revenue,
        donor_units_sold=float(donor_sales["units_sold"] or 0),
        donor_sale_count=int(donor_sales["sale_count"] or 0),
        donor_profit=donor_profit,
        total_revenue=total_revenue,
        recovery_pct=recovery_pct,
        remaining_to_break_even=remaining_to_break_even,
        break_even_status=break_even_status,
    )
    valuation_button = f"<a href=\"{url_for('vehicle_valuation', vehicle_id=vehicle_id)}\" style=\"position:fixed;right:22px;bottom:22px;z-index:9998;background:#15803d;color:white;padding:14px 18px;border-radius:12px;text-decoration:none;font-weight:800\">Market Valuation &amp; Deal Score</a>"
    return vehicle_page.replace("</body>", valuation_button + "</body>")

@app.route("/vehicles/<int:vehicle_id>/documents", methods=["POST"])
@login_required
def vehicle_document_add(vehicle_id):
    upload = request.files.get("document")
    try:
        filename = save_upload(upload)
        if not filename:
            raise ValueError("Choose a document to upload.")
        conn = db()
        conn.execute(
            """INSERT INTO vehicle_documents(vehicle_id,document_type,filename,description)
               VALUES(?,?,?,?)""",
            (
                vehicle_id,
                request.form.get("document_type") or "Other",
                filename,
                request.form.get("description"),
            ),
        )
        conn.commit()
        conn.close()
        log_action("Document uploaded", "vehicle", vehicle_id, request.form.get("document_type") or "Other")
        flash("Document uploaded.", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id) + "#documents")

@app.route("/vehicles/<int:vehicle_id>/documents/<int:document_id>/delete", methods=["POST"])
@login_required
def vehicle_document_delete(vehicle_id, document_id):
    conn = db()
    document = conn.execute(
        "SELECT * FROM vehicle_documents WHERE id=? AND vehicle_id=?",
        (document_id, vehicle_id),
    ).fetchone()
    if document:
        conn.execute("DELETE FROM vehicle_documents WHERE id=?", (document_id,))
        conn.commit()
    conn.close()
    if document:
        try:
            (UPLOAD_DIR / document["filename"]).unlink(missing_ok=True)
        except OSError:
            pass
        log_action("Document deleted", "vehicle", vehicle_id, document["document_type"])
        flash("Document deleted.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id) + "#documents")

@app.route("/vehicles/<int:vehicle_id>/expense", methods=["POST"])
@login_required
def expense_add(vehicle_id):
    try:
        receipt = save_upload(request.files.get("receipt"))
        cost = float(request.form.get("cost_inc_gst") or 0)
        gst = round(cost / 11, 2)
        conn = db()
        conn.execute("""
            INSERT INTO expenses(vehicle_id,expense_date,category,description,supplier,paid_by,
                                 cost_inc_gst,gst_amount,receipt_filename,notes)
            VALUES(?,?,?,?,?,?,?,?,?,?)
        """, (vehicle_id, request.form.get("expense_date"), request.form["category"],
              request.form["description"], request.form.get("supplier"),
              request.form["paid_by"], cost, gst, receipt, request.form.get("notes")))
        conn.commit()
        conn.close()
        flash("Expense added.", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))

@app.route("/vehicles/<int:vehicle_id>/photos", methods=["POST"])
@login_required
def photo_add(vehicle_id):
    files = request.files.getlist("photos")
    caption = request.form.get("caption")
    saved = 0
    conn = db()
    try:
        for file in files:
            filename = save_upload(file)
            if filename:
                conn.execute(
                    "INSERT INTO vehicle_photos(vehicle_id,filename,caption) VALUES(?,?,?)",
                    (vehicle_id, filename, caption),
                )
                saved += 1
        conn.commit()
        flash(f"{saved} photo(s) added.", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    finally:
        conn.close()
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/job-card", methods=["POST"])
@login_required
def job_card_add(vehicle_id):
    actual = float(request.form.get("actual_cost_inc_gst") or 0)
    gst = round(actual / 11, 2)
    conn = db()
    cursor = conn.execute("""
        INSERT INTO job_cards(
            vehicle_id,job_date,category,description,supplier,paid_by,
            estimated_cost,actual_cost_inc_gst,gst_amount,status,notes,
            labour_operation_id,labour_hours,labour_rate,labour_source,labour_code,procedure_url
        ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        vehicle_id,
        request.form.get("job_date"),
        request.form.get("category"),
        request.form.get("description"),
        request.form.get("supplier"),
        request.form.get("paid_by"),
        float(request.form.get("estimated_cost") or 0),
        actual,
        gst,
        request.form.get("status") or "Open",
        request.form.get("notes"),
        int(request.form.get("labour_operation_id")) if request.form.get("labour_operation_id") else None,
        float(request.form.get("labour_hours") or 0),
        float(request.form.get("labour_rate") or 0),
        request.form.get("labour_source"),
        request.form.get("labour_code"),
        request.form.get("procedure_url"),
    ))
    job_id = cursor.lastrowid
    conn.execute("""
        INSERT INTO job_card_history(job_card_id,old_status,new_status,changed_by,note)
        VALUES(?,?,?,?,?)
    """, (
        job_id, None, request.form.get("status") or "Open",
        session.get("display_name"), "Job card created"
    ))
    conn.commit()
    conn.close()
    flash("Job card added.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))



@app.route("/job-cards/<int:job_id>/edit", methods=["GET", "POST"])
@login_required
def job_card_edit(job_id):
    conn = db()
    job = conn.execute("SELECT * FROM job_cards WHERE id=?", (job_id,)).fetchone()
    if not job:
        conn.close()
        return "Job card not found", 404

    if request.method == "POST":
        old_status = job["status"]
        new_status = request.form.get("status") or old_status
        actual = float(request.form.get("actual_cost_inc_gst") or 0)
        gst = round(actual / 11, 2)

        conn.execute("""
            UPDATE job_cards
            SET job_date=?, category=?, description=?, supplier=?, paid_by=?,
                estimated_cost=?, actual_cost_inc_gst=?, gst_amount=?, status=?, notes=?,
                labour_hours=?,labour_rate=?,labour_source=?,labour_code=?,procedure_url=?
            WHERE id=?
        """, (
            request.form.get("job_date"),
            request.form.get("category"),
            request.form.get("description"),
            request.form.get("supplier"),
            request.form.get("paid_by"),
            float(request.form.get("estimated_cost") or 0),
            actual,
            gst,
            new_status,
            request.form.get("notes"),
            float(request.form.get("labour_hours") or 0),
            float(request.form.get("labour_rate") or 0),
            request.form.get("labour_source"),
            request.form.get("labour_code"),
            request.form.get("procedure_url"),
            job_id,
        ))

        change_note = request.form.get("change_note") or "Job card updated"
        conn.execute("""
            INSERT INTO job_card_history(job_card_id,old_status,new_status,changed_by,note)
            VALUES(?,?,?,?,?)
        """, (
            job_id, old_status, new_status,
            session.get("display_name"), change_note
        ))
        conn.commit()
        vehicle_id = job["vehicle_id"]
        conn.close()
        flash("Job card updated.", "success")
        return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))

    history = conn.execute(
        "SELECT * FROM job_card_history WHERE job_card_id=? ORDER BY changed_at DESC,id DESC",
        (job_id,)
    ).fetchall()
    conn.close()
    return render_template("job_card_edit.html", job=job, history=history)


@app.route("/job-cards/<int:job_id>/reopen", methods=["POST"])
@login_required
def job_card_reopen(job_id):
    conn = db()
    job = conn.execute("SELECT * FROM job_cards WHERE id=?", (job_id,)).fetchone()
    if not job:
        conn.close()
        return "Job card not found", 404

    old_status = job["status"]
    new_status = request.form.get("status") or "Open"
    note = request.form.get("note") or "Job card reopened"

    conn.execute("UPDATE job_cards SET status=? WHERE id=?", (new_status, job_id))
    conn.execute("""
        INSERT INTO job_card_history(job_card_id,old_status,new_status,changed_by,note)
        VALUES(?,?,?,?,?)
    """, (
        job_id, old_status, new_status,
        session.get("display_name"), note
    ))
    conn.commit()
    vehicle_id = job["vehicle_id"]
    conn.close()
    flash(f"Job card reopened as {new_status}.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/sale", methods=["POST"])
@login_required
def sale_add(vehicle_id):
    price = float(request.form.get("sale_price_inc_gst") or 0)
    gst = round(price / 11, 2)
    conn = db()
    existing = conn.execute("SELECT * FROM sales WHERE vehicle_id=?", (vehicle_id,)).fetchone()
    invoice_number = existing["invoice_number"] if existing and existing["invoice_number"] else next_invoice_number(conn)
    contract_number = existing["contract_number"] if existing and existing["contract_number"] else next_contract_number(conn)

    conn.execute("""
        INSERT INTO sales(
            vehicle_id,sale_date,buyer_name,buyer_phone,buyer_email,buyer_address,
            sale_price_inc_gst,sale_gst,advertising_cost,transfer_cost,invoice_number,
            deposit_amount,deposit_date,payment_method,trade_in_description,trade_in_value,
            warranty_type,warranty_expiry,contract_number,notes
        )
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(vehicle_id) DO UPDATE SET
          sale_date=excluded.sale_date,
          buyer_name=excluded.buyer_name,
          buyer_phone=excluded.buyer_phone,
          buyer_email=excluded.buyer_email,
          buyer_address=excluded.buyer_address,
          sale_price_inc_gst=excluded.sale_price_inc_gst,
          sale_gst=excluded.sale_gst,
          advertising_cost=excluded.advertising_cost,
          transfer_cost=excluded.transfer_cost,
          invoice_number=COALESCE(sales.invoice_number,excluded.invoice_number),
          deposit_amount=excluded.deposit_amount,
          deposit_date=excluded.deposit_date,
          payment_method=excluded.payment_method,
          trade_in_description=excluded.trade_in_description,
          trade_in_value=excluded.trade_in_value,
          warranty_type=excluded.warranty_type,
          warranty_expiry=excluded.warranty_expiry,
          contract_number=COALESCE(sales.contract_number,excluded.contract_number),
          notes=excluded.notes
    """, (
        vehicle_id,
        request.form.get("sale_date"),
        request.form.get("buyer_name"),
        request.form.get("buyer_phone"),
        request.form.get("buyer_email"),
        request.form.get("buyer_address"),
        price,
        gst,
        float(request.form.get("advertising_cost") or 0),
        float(request.form.get("transfer_cost") or 0),
        invoice_number,
        float(request.form.get("deposit_amount") or 0),
        request.form.get("deposit_date"),
        request.form.get("payment_method"),
        request.form.get("trade_in_description"),
        float(request.form.get("trade_in_value") or 0),
        request.form.get("warranty_type"),
        request.form.get("warranty_expiry"),
        contract_number,
        request.form.get("notes"),
    ))
    conn.execute("UPDATE vehicles SET status='Sold' WHERE id=?", (vehicle_id,))
    conn.commit()
    conn.close()
    log_action("Sale recorded", "vehicle", vehicle_id, f"{invoice_number} / {contract_number}")
    flash(f"Sale recorded. Invoice {invoice_number} and contract {contract_number} created.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.route("/import-excel", methods=["GET", "POST"])
@login_required
def import_excel():
    result = None
    if request.method == "POST":
        uploaded = request.files.get("excel_file")
        if not uploaded or not uploaded.filename.lower().endswith(".xlsx"):
            flash("Please choose an .xlsx Excel workbook.", "error")
            return redirect(url_for("import_excel"))

        temp_path = BASE_DIR / f"import_{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{secure_filename(uploaded.filename)}"
        uploaded.save(temp_path)
        imported = 0
        skipped = 0
        warnings = []
        workbook = None

        try:
            workbook = load_workbook(temp_path, data_only=True)
            sheet = None
            for candidate in ["Vehicles", "Vehicle Inventory"]:
                if candidate in workbook.sheetnames:
                    sheet = workbook[candidate]
                    break
            if sheet is None:
                raise ValueError("The workbook needs a sheet named Vehicles or Vehicle Inventory.")

            headers = {}
            header_row = None
            for row_number in range(1, min(sheet.max_row, 15) + 1):
                values = [sheet.cell(row_number, col).value for col in range(1, sheet.max_column + 1)]
                if any(str(v).strip() in ("Stock No.", "Stock Number") for v in values if v is not None):
                    header_row = row_number
                    headers = {
                        str(sheet.cell(row_number, col).value).strip(): col
                        for col in range(1, sheet.max_column + 1)
                        if sheet.cell(row_number, col).value is not None
                    }
                    break
            if not header_row:
                raise ValueError("Could not find the vehicle heading row.")

            def value(row, *names):
                for name in names:
                    col = headers.get(name)
                    if col:
                        return sheet.cell(row, col).value
                return None

            conn = db()
            for row in range(header_row + 1, sheet.max_row + 1):
                make = str(value(row, "Make") or "").strip()
                model = str(value(row, "Model") or "").strip()
                if not make and not model:
                    continue

                stock = str(value(row, "Stock No.", "Stock Number") or "").strip()
                if not stock:
                    stock = next_stock_number(conn)

                existing = conn.execute(
                    "SELECT id FROM vehicles WHERE stock_no=?",
                    (stock,)
                ).fetchone()
                if existing:
                    skipped += 1
                    continue

                vin = str(value(row, "VIN") or "").strip() or None
                if vin and conn.execute("SELECT id FROM vehicles WHERE vin=?", (vin,)).fetchone():
                    warnings.append(f"{stock}: duplicate VIN skipped")
                    skipped += 1
                    continue

                price = clean_number(value(row, "Purchase Price Inc GST", "Purchase Price"))
                gst = clean_number(value(row, "Purchase GST"))
                if not gst and price:
                    gst = round(price / 11, 2)

                purchase_date = value(row, "Purchase Date")
                rego_expiry = value(row, "Registration Expiry", "Rego Expiry")
                for_date = lambda x: x.date().isoformat() if hasattr(x, "date") else (str(x) if x else None)

                conn.execute("""
                    INSERT INTO vehicles(
                        stock_no,status,purchase_date,make,model,variant,year,vin,registration,
                        odometer_km,colour,purchase_price_inc_gst,purchase_gst,
                        barry_contribution,matt_contribution,rego_expiry,photo_filename,notes
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    stock,
                    value(row, "Status") or "In Stock",
                    for_date(purchase_date),
                    make,
                    model,
                    value(row, "Variant"),
                    int(clean_number(value(row, "Year"))) or None,
                    vin,
                    value(row, "Registration"),
                    int(clean_number(value(row, "Odometer km", "Odometer"))) or None,
                    value(row, "Colour"),
                    price,
                    gst,
                    clean_number(value(row, "Barry Contribution", "Barry Purchase Contribution")),
                    clean_number(value(row, "Matt Contribution", "Matt Purchase Contribution")),
                    for_date(rego_expiry),
                    value(row, "Main Photo Link"),
                    value(row, "Notes"),
                ))
                imported += 1

            conn.commit()
            conn.close()
            result = {"imported": imported, "skipped": skipped, "warnings": warnings}
            flash(f"Import completed: {imported} vehicle(s) added, {skipped} skipped.", "success")
        except Exception as exc:
            flash(f"Import failed: {exc}", "error")
        finally:
            if workbook is not None:
                try:
                    workbook.close()
                except Exception:
                    pass
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except PermissionError:
                    # Windows may briefly retain the file handle. The next import
                    # uses a new temporary name, so leaving this temp copy is safe.
                    pass

    return render_template("import_excel.html", result=result)


@app.route("/vehicles/<int:vehicle_id>/report")
@login_required
def vehicle_report(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    expenses = conn.execute("SELECT * FROM expenses WHERE vehicle_id=? ORDER BY expense_date,id", (vehicle_id,)).fetchall()
    jobs = conn.execute("SELECT * FROM job_cards WHERE vehicle_id=? ORDER BY job_date,id", (vehicle_id,)).fetchall()
    sale = conn.execute("SELECT * FROM sales WHERE vehicle_id=?", (vehicle_id,)).fetchone()
    conn.close()
    if not vehicle:
        return "Vehicle not found", 404
    expense_total = sum(x["cost_inc_gst"] for x in expenses)
    job_total = sum((x["actual_cost_inc_gst"] if x["actual_cost_inc_gst"] > 0 else x["estimated_cost"]) for x in jobs)
    sale_price = sale["sale_price_inc_gst"] if sale else 0
    selling_costs = (sale["advertising_cost"] + sale["transfer_cost"]) if sale else 0
    profit = sale_price - vehicle["purchase_price_inc_gst"] - expense_total - job_total - selling_costs
    return render_template("vehicle_report.html", vehicle=vehicle, expenses=expenses, jobs=jobs,
                           sale=sale, expense_total=expense_total, job_total=job_total, profit=profit)

@app.route("/backup")
@login_required
def backup_database():
    return redirect(url_for("backup_centre"))


def create_full_backup(prefix="bam_full_backup"):
    BACKUP_DIR.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"{prefix}_{stamp}.zip"
    temp_db = BACKUP_DIR / f".snapshot_{stamp}.db"
    source = sqlite3.connect(DB_PATH)
    target = sqlite3.connect(temp_db)
    with target:
        source.backup(target)
    source.close(); target.close()
    with zipfile.ZipFile(backup_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(temp_db, "database/bam_motor_group.db")
        for folder_name in ("uploads", "reports"):
            folder = PROJECT_DIR / folder_name
            if folder.exists():
                for file in folder.rglob("*"):
                    if file.is_file():
                        zf.write(file, file.relative_to(PROJECT_DIR))
        zf.writestr("BACKUP_INFO.txt", f"{APP_NAME} v{APP_VERSION}\nCreated: {datetime.now().isoformat()}\n")
    temp_db.unlink(missing_ok=True)
    return backup_path


@app.route("/backups")
@login_required
def backup_centre():
    backups = sorted(BACKUP_DIR.glob("*.zip"), key=lambda x: x.stat().st_mtime, reverse=True)
    rows = [{"name": b.name, "size": b.stat().st_size, "modified": datetime.fromtimestamp(b.stat().st_mtime)} for b in backups]
    return render_template("backup_centre.html", backups=rows)


@app.route("/backups/create", methods=["POST"])
@login_required
def backup_create():
    path = create_full_backup()
    log_action("Full backup created", "system", None, path.name)
    return send_file(path, as_attachment=True, download_name=path.name)


@app.route("/backups/download/<path:filename>")
@login_required
def backup_download(filename):
    safe = Path(filename).name
    path = BACKUP_DIR / safe
    if not path.exists() or path.suffix.lower() != ".zip":
        return "Backup not found", 404
    return send_file(path, as_attachment=True, download_name=path.name)


@app.route("/backups/upload", methods=["POST"])
@owner_required
def backup_upload():
    file = request.files.get("backup_file")
    if not file or not file.filename or not file.filename.lower().endswith(".zip"):
        flash("Choose a BAM ZIP backup.", "error")
        return redirect(url_for("backup_centre"))
    name = secure_filename(file.filename)
    destination = BACKUP_DIR / name
    file.save(destination)
    try:
        with zipfile.ZipFile(destination) as zf:
            if "database/bam_motor_group.db" not in zf.namelist():
                raise ValueError("This is not a valid BAM full backup.")
    except Exception as exc:
        destination.unlink(missing_ok=True)
        flash(f"Backup upload failed: {exc}", "error")
        return redirect(url_for("backup_centre"))
    flash("Backup uploaded and ready to restore.", "success")
    return redirect(url_for("backup_centre"))


@app.route("/backups/restore", methods=["POST"])
@owner_required
def backup_restore():
    filename = Path(request.form.get("filename") or "").name
    confirmation = (request.form.get("confirmation") or "").strip().upper()
    if confirmation != "RESTORE":
        flash("Type RESTORE to confirm.", "error")
        return redirect(url_for("backup_centre"))
    path = BACKUP_DIR / filename
    if not path.exists() or path.suffix.lower() != ".zip":
        flash("Backup file not found.", "error")
        return redirect(url_for("backup_centre"))
    create_full_backup("pre_restore_backup")
    restore_dir = BACKUP_DIR / ".restore_temp"
    if restore_dir.exists(): shutil.rmtree(restore_dir)
    restore_dir.mkdir()
    try:
        with zipfile.ZipFile(path) as zf:
            for member in zf.infolist():
                target = (restore_dir / member.filename).resolve()
                if restore_dir.resolve() not in target.parents and target != restore_dir.resolve():
                    raise ValueError("Unsafe backup archive.")
            zf.extractall(restore_dir)
        restored_db = restore_dir / "database" / "bam_motor_group.db"
        if not restored_db.exists(): raise ValueError("The backup does not contain a database.")
        shutil.copy2(restored_db, DB_PATH)
        for folder_name in ("uploads", "reports"):
            source_folder = restore_dir / folder_name
            if source_folder.exists():
                destination = PROJECT_DIR / folder_name
                destination.mkdir(exist_ok=True)
                for file in source_folder.rglob("*"):
                    if file.is_file():
                        out = destination / file.relative_to(source_folder)
                        out.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(file, out)
        init_db()
        session.clear()
        flash("Backup restored successfully. Please sign in again.", "success")
        return redirect(url_for("login"))
    except Exception as exc:
        flash(f"Restore failed: {exc}", "error")
        return redirect(url_for("backup_centre"))
    finally:
        if restore_dir.exists(): shutil.rmtree(restore_dir)


@app.route("/contacts", methods=["GET", "POST"])
@login_required
def contacts_page():
    conn = db()
    if request.method == "POST":
        conn.execute("""
            INSERT INTO contacts(contact_type,name,company,phone,email,address,licence_no,crm_status,source,tags,preferred_contact,next_follow_up_date,notes)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            request.form.get("contact_type"),
            request.form.get("name"),
            request.form.get("company"),
            request.form.get("phone"),
            request.form.get("email"),
            request.form.get("address"),
            request.form.get("licence_no"),
            request.form.get("crm_status") or "Active",
            request.form.get("source"),
            request.form.get("tags"),
            request.form.get("preferred_contact"),
            request.form.get("next_follow_up_date"),
            request.form.get("notes"),
        ))
        conn.commit()
        flash("Contact saved.", "success")
        return redirect(url_for("contacts_page"))

    q = request.args.get("q", "").strip()
    if q:
        rows = conn.execute("""
            SELECT * FROM contacts
            WHERE name LIKE ? OR phone LIKE ? OR email LIKE ? OR contact_type LIKE ?
            ORDER BY name
        """, tuple([f"%{q}%"] * 4)).fetchall()
    else:
        rows = conn.execute("SELECT * FROM contacts ORDER BY name").fetchall()
    conn.close()
    return render_template("contacts.html", contacts=rows, q=q)


@app.route("/reports/profit")
@login_required
def profit_report():
    conn = db()
    rows = conn.execute("""
        SELECT v.id,v.stock_no,v.year,v.make,v.model,v.status,
               v.purchase_price_inc_gst,
               COALESCE((SELECT SUM(e.cost_inc_gst) FROM expenses e WHERE e.vehicle_id=v.id),0) AS expenses,
               COALESCE((SELECT SUM(CASE WHEN j.actual_cost_inc_gst>0 THEN j.actual_cost_inc_gst ELSE j.estimated_cost END)
                         FROM job_cards j WHERE j.vehicle_id=v.id),0) AS jobs,
               COALESCE((SELECT s.sale_price_inc_gst FROM sales s WHERE s.vehicle_id=v.id),0) AS sale_price,
               COALESCE((SELECT s.advertising_cost+s.transfer_cost FROM sales s WHERE s.vehicle_id=v.id),0) AS selling_costs
        FROM vehicles v
        ORDER BY v.id DESC
    """).fetchall()
    conn.close()
    report_rows = []
    for row in rows:
        total_cost = row["purchase_price_inc_gst"] + row["expenses"] + row["jobs"] + row["selling_costs"]
        profit = row["sale_price"] - total_cost
        report_rows.append(dict(row) | {"total_cost": total_cost, "profit": profit})
    totals = {
        "purchase": sum(r["purchase_price_inc_gst"] for r in report_rows),
        "expenses": sum(r["expenses"] for r in report_rows),
        "jobs": sum(r["jobs"] for r in report_rows),
        "sales": sum(r["sale_price"] for r in report_rows),
        "profit": sum(r["profit"] for r in report_rows),
    }
    return render_template("profit_report.html", rows=report_rows, totals=totals)


@app.route("/reports/dismantling-profit")
@login_required
def dismantling_profit_report():
    conn = db()
    rows = conn.execute("""
        SELECT v.id,v.stock_no,v.year,v.make,v.model,v.status,v.dismantling_status,
               v.purchase_price_inc_gst,v.shell_sale_price,
               COALESCE((SELECT SUM(e.cost_inc_gst) FROM expenses e WHERE e.vehicle_id=v.id),0) AS expenses,
               COALESCE((SELECT SUM(CASE WHEN j.actual_cost_inc_gst>0 THEN j.actual_cost_inc_gst ELSE j.estimated_cost END)
                         FROM job_cards j WHERE j.vehicle_id=v.id),0) AS jobs,
               COALESCE((SELECT SUM(se.cost_inc_gst) FROM service_entries se WHERE se.vehicle_id=v.id),0) AS services,
               COALESCE((SELECT SUM(pu.quantity_used*pu.unit_cost_inc_gst) FROM part_usage pu WHERE pu.vehicle_id=v.id),0) AS parts_used,
               COALESCE((SELECT SUM(ps.sale_price) FROM part_sales ps JOIN parts p ON p.id=ps.part_id WHERE p.vehicle_id=v.id),0) AS parts_revenue,
               COALESCE((SELECT SUM(ps.quantity) FROM part_sales ps JOIN parts p ON p.id=ps.part_id WHERE p.vehicle_id=v.id),0) AS units_sold,
               COALESCE((SELECT SUM(p.quantity_on_hand*p.selling_price) FROM parts p WHERE p.vehicle_id=v.id),0) AS remaining_retail,
               COALESCE((SELECT COUNT(*) FROM parts p WHERE p.vehicle_id=v.id),0) AS part_lines,
               COALESCE((SELECT s.sale_price_inc_gst FROM sales s WHERE s.vehicle_id=v.id),0) AS vehicle_sale
        FROM vehicles v
        WHERE COALESCE(v.vehicle_purpose,'Retail Sale')='Parts Vehicle' OR v.status='BER'
        ORDER BY v.id DESC
    """).fetchall()
    conn.close()
    report_rows = []
    for row in rows:
        item = dict(row)
        invested = sum(float(item[k] or 0) for k in ["purchase_price_inc_gst","expenses","jobs","services","parts_used"])
        revenue = float(item["parts_revenue"] or 0) + float(item["shell_sale_price"] or 0) + float(item["vehicle_sale"] or 0)
        item["invested"] = invested
        item["total_revenue"] = revenue
        item["profit"] = revenue - invested
        item["recovery_pct"] = (revenue / invested * 100) if invested > 0 else (100.0 if revenue > 0 else 0.0)
        item["remaining_to_break_even"] = max(0.0, invested - revenue)
        report_rows.append(item)
    totals = {
        "invested": sum(r["invested"] for r in report_rows),
        "parts_revenue": sum(float(r["parts_revenue"] or 0) for r in report_rows),
        "shell_revenue": sum(float(r["shell_sale_price"] or 0) for r in report_rows),
        "remaining_retail": sum(float(r["remaining_retail"] or 0) for r in report_rows),
        "profit": sum(r["profit"] for r in report_rows),
    }
    return render_template("dismantling_profit_report.html", rows=report_rows, totals=totals)


@app.route("/invoices")
@login_required
def invoice_centre():
    q = (request.args.get("q") or "").strip()
    status = (request.args.get("status") or "").strip()
    conn = db()
    sql = """
        SELECT s.*, v.stock_no, v.asset_type, v.year, v.make, v.model, v.registration
        FROM sales s JOIN vehicles v ON v.id=s.vehicle_id
        WHERE 1=1
    """
    params = []
    if q:
        like = f"%{q}%"
        sql += " AND (s.invoice_number LIKE ? OR s.buyer_name LIKE ? OR v.stock_no LIKE ? OR v.registration LIKE ?)"
        params.extend([like, like, like, like])
    if status:
        sql += " AND COALESCE(s.invoice_status,'Draft')=?"
        params.append(status)
    sql += " ORDER BY COALESCE(s.sale_date, s.id) DESC, s.id DESC"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return render_template("invoice_centre.html", invoices=rows, q=q, status=status)


@app.route("/vehicles/<int:vehicle_id>/invoice/edit", methods=["GET", "POST"])
@login_required
def sale_invoice_edit(vehicle_id):
    import json
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    sale = conn.execute("SELECT * FROM sales WHERE vehicle_id=?", (vehicle_id,)).fetchone()
    if not vehicle or not sale:
        conn.close()
        return "Sale invoice is not available until a sale is recorded.", 404

    if request.method == "POST":
        tracked = [
            "sale_date", "buyer_name", "buyer_phone", "buyer_email", "buyer_address",
            "sale_price_inc_gst", "advertising_cost", "transfer_cost", "deposit_amount",
            "deposit_date", "payment_method", "trade_in_description", "trade_in_value",
            "warranty_type", "warranty_expiry", "invoice_status", "notes"
        ]
        old_values = {key: sale[key] for key in tracked}
        price = float(request.form.get("sale_price_inc_gst") or 0)
        new_values = {
            "sale_date": request.form.get("sale_date"),
            "buyer_name": request.form.get("buyer_name"),
            "buyer_phone": request.form.get("buyer_phone"),
            "buyer_email": request.form.get("buyer_email"),
            "buyer_address": request.form.get("buyer_address"),
            "sale_price_inc_gst": price,
            "advertising_cost": float(request.form.get("advertising_cost") or 0),
            "transfer_cost": float(request.form.get("transfer_cost") or 0),
            "deposit_amount": float(request.form.get("deposit_amount") or 0),
            "deposit_date": request.form.get("deposit_date"),
            "payment_method": request.form.get("payment_method"),
            "trade_in_description": request.form.get("trade_in_description"),
            "trade_in_value": float(request.form.get("trade_in_value") or 0),
            "warranty_type": request.form.get("warranty_type"),
            "warranty_expiry": request.form.get("warranty_expiry"),
            "invoice_status": request.form.get("invoice_status") or "Draft",
            "notes": request.form.get("notes"),
        }
        gst = round(price / 11, 2)
        conn.execute("""
            UPDATE sales SET
              sale_date=?,buyer_name=?,buyer_phone=?,buyer_email=?,buyer_address=?,
              sale_price_inc_gst=?,sale_gst=?,advertising_cost=?,transfer_cost=?,
              deposit_amount=?,deposit_date=?,payment_method=?,trade_in_description=?,trade_in_value=?,
              warranty_type=?,warranty_expiry=?,invoice_status=?,notes=?,updated_at=CURRENT_TIMESTAMP
            WHERE vehicle_id=?
        """, (
            new_values["sale_date"], new_values["buyer_name"], new_values["buyer_phone"],
            new_values["buyer_email"], new_values["buyer_address"], price, gst,
            new_values["advertising_cost"], new_values["transfer_cost"], new_values["deposit_amount"],
            new_values["deposit_date"], new_values["payment_method"], new_values["trade_in_description"],
            new_values["trade_in_value"], new_values["warranty_type"], new_values["warranty_expiry"],
            new_values["invoice_status"], new_values["notes"], vehicle_id
        ))
        changed = {k: {"from": old_values.get(k), "to": new_values.get(k)} for k in tracked if str(old_values.get(k) or "") != str(new_values.get(k) or "")}
        note = request.form.get("change_note") or "Invoice updated"
        conn.execute("""
            INSERT INTO invoice_history(sale_id,vehicle_id,invoice_number,changed_by,change_note,old_values,new_values)
            VALUES(?,?,?,?,?,?,?)
        """, (sale["id"], vehicle_id, sale["invoice_number"], session.get("display_name"), note,
              json.dumps(old_values, default=str), json.dumps(new_values, default=str)))
        conn.commit()
        conn.close()
        log_action("Invoice edited", "vehicle", vehicle_id, f"{sale['invoice_number']}: {note}; {len(changed)} field(s) changed")
        flash(f"Invoice {sale['invoice_number']} updated.", "success")
        if request.form.get("save_continue"):
            return redirect(url_for("sale_invoice_edit", vehicle_id=vehicle_id))
        return redirect(url_for("sale_invoice", vehicle_id=vehicle_id))

    history = conn.execute("SELECT * FROM invoice_history WHERE sale_id=? ORDER BY changed_at DESC,id DESC", (sale["id"],)).fetchall()
    conn.close()
    return render_template("invoice_edit.html", vehicle=vehicle, sale=sale, history=history)


@app.route("/vehicles/<int:vehicle_id>/deal-file", methods=["GET", "POST"])
@login_required
def deal_file(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    sale = conn.execute("SELECT * FROM sales WHERE vehicle_id=?", (vehicle_id,)).fetchone()
    if not vehicle or not sale:
        conn.close()
        return "A deal file is available after a sale has been recorded.", 404

    if request.method == "POST":
        conn.execute("""
            UPDATE sales SET delivery_status=?,delivery_date=?,keys_handed_over=?,
              registration_transferred=?,customer_signature_received=?,
              finance_documents_complete=?,warranty_documents_complete=?,deal_notes=?,
              updated_at=CURRENT_TIMESTAMP
            WHERE vehicle_id=?
        """, (
            request.form.get("delivery_status") or "Preparing",
            request.form.get("delivery_date") or None,
            1 if request.form.get("keys_handed_over") else 0,
            1 if request.form.get("registration_transferred") else 0,
            1 if request.form.get("customer_signature_received") else 0,
            1 if request.form.get("finance_documents_complete") else 0,
            1 if request.form.get("warranty_documents_complete") else 0,
            request.form.get("deal_notes"),
            vehicle_id,
        ))
        conn.commit()
        conn.close()
        log_action("Deal file updated", "vehicle", vehicle_id, request.form.get("delivery_status") or "Preparing")
        flash("Deal file updated.", "success")
        return redirect(url_for("deal_file", vehicle_id=vehicle_id))

    documents = conn.execute(
        "SELECT * FROM vehicle_documents WHERE vehicle_id=? ORDER BY uploaded_at DESC,id DESC",
        (vehicle_id,),
    ).fetchall()
    photos = conn.execute(
        "SELECT * FROM vehicle_photos WHERE vehicle_id=? ORDER BY id DESC", (vehicle_id,)
    ).fetchall()
    history = conn.execute(
        "SELECT * FROM invoice_history WHERE sale_id=? ORDER BY changed_at DESC,id DESC", (sale["id"],)
    ).fetchall()
    expenses = conn.execute(
        "SELECT * FROM expenses WHERE vehicle_id=? ORDER BY expense_date DESC,id DESC", (vehicle_id,)
    ).fetchall()
    jobs = conn.execute(
        "SELECT * FROM job_cards WHERE vehicle_id=? ORDER BY job_date DESC,id DESC", (vehicle_id,)
    ).fetchall()
    conn.close()
    balance_due = max(float(sale["sale_price_inc_gst"] or 0) - float(sale["deposit_amount"] or 0), 0)
    return render_template(
        "deal_file.html", vehicle=vehicle, sale=sale, documents=documents, photos=photos,
        history=history, expenses=expenses, jobs=jobs, balance_due=balance_due
    )


@app.route("/vehicles/<int:vehicle_id>/invoice")
@login_required
def sale_invoice(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    sale = conn.execute("SELECT * FROM sales WHERE vehicle_id=?", (vehicle_id,)).fetchone()
    conn.close()
    if not vehicle or not sale:
        return "Sale invoice is not available until a sale is recorded.", 404
    sale_ex_gst = sale["sale_price_inc_gst"] - sale["sale_gst"]
    return render_template("sale_invoice.html", vehicle=vehicle, sale=sale, sale_ex_gst=sale_ex_gst)


@app.route("/export/vehicles.csv")
@login_required
def export_vehicles_csv():
    conn = db()
    rows = conn.execute("""
        SELECT stock_no,status,purchase_date,year,make,model,variant,vin,registration,
               odometer_km,purchase_price_inc_gst,purchase_gst,barry_contribution,
               matt_contribution,rego_expiry,ppsr_number,roadworthy_status,service_due_date
        FROM vehicles ORDER BY id
    """).fetchall()
    conn.close()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(rows[0].keys() if rows else [
        "stock_no","status","purchase_date","year","make","model","variant","vin",
        "registration","odometer_km","purchase_price_inc_gst","purchase_gst",
        "barry_contribution","matt_contribution","rego_expiry","ppsr_number",
        "roadworthy_status","service_due_date"
    ])
    for row in rows:
        writer.writerow(tuple(row))
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=BAM_vehicle_export.csv"},
    )


@app.route("/export/profit.csv")
@login_required
def export_profit_csv():
    conn = db()
    rows = conn.execute("""
        SELECT v.stock_no,v.year,v.make,v.model,v.status,
               v.purchase_price_inc_gst,
               COALESCE((SELECT SUM(e.cost_inc_gst) FROM expenses e WHERE e.vehicle_id=v.id),0) AS expenses,
               COALESCE((SELECT SUM(CASE WHEN j.actual_cost_inc_gst>0 THEN j.actual_cost_inc_gst ELSE j.estimated_cost END)
                         FROM job_cards j WHERE j.vehicle_id=v.id),0) AS jobs,
               COALESCE((SELECT s.sale_price_inc_gst FROM sales s WHERE s.vehicle_id=v.id),0) AS sale_price
        FROM vehicles v ORDER BY v.id
    """).fetchall()
    conn.close()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Stock","Year","Make","Model","Status","Purchase","Expenses","Jobs","Sale","Profit/Loss"])
    for row in rows:
        profit = row["sale_price"] - row["purchase_price_inc_gst"] - row["expenses"] - row["jobs"]
        writer.writerow([
            row["stock_no"],row["year"],row["make"],row["model"],row["status"],
            row["purchase_price_inc_gst"],row["expenses"],row["jobs"],row["sale_price"],profit
        ])
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=BAM_profit_export.csv"},
    )


@app.route("/vehicles/<int:vehicle_id>/dismantling", methods=["GET", "POST"])
@login_required
def vehicle_dismantling(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close()
        return "Vehicle not found", 404

    if request.method == "POST":
        action = request.form.get("action") or "add_part"
        if action == "update_status":
            status = request.form.get("dismantling_status") or "Not Started"
            conn.execute(
                "UPDATE vehicles SET vehicle_purpose='Parts Vehicle', dismantling_status=? WHERE id=?",
                (status, vehicle_id),
            )
            conn.commit()
            conn.close()
            flash("Dismantling status updated.", "success")
            return redirect(url_for("vehicle_dismantling", vehicle_id=vehicle_id))

        if action == "shell_sale":
            price = float(request.form.get("shell_sale_price") or 0)
            conn.execute(
                "UPDATE vehicles SET vehicle_purpose='Parts Vehicle', shell_sale_price=?, shell_sale_date=? WHERE id=?",
                (price, request.form.get("shell_sale_date") or None, vehicle_id),
            )
            conn.commit()
            conn.close()
            flash("Shell sale recorded.", "success")
            return redirect(url_for("vehicle_dismantling", vehicle_id=vehicle_id))

        part_name = (request.form.get("part_name") or "").strip()
        if not part_name:
            conn.close()
            flash("Part name is required.", "error")
            return redirect(url_for("vehicle_dismantling", vehicle_id=vehicle_id))
        qty = float(request.form.get("quantity_on_hand") or 1)
        cost = float(request.form.get("unit_cost_inc_gst") or 0)
        selling = float(request.form.get("selling_price") or 0)
        gst = round(cost / 11, 2)
        # PRT numbers are allocated by the server from one business-wide sequence.
        # Ignore stale PRT values posted by an open browser form so a duplicate SKU
        # can never be reused accidentally. Custom non-PRT numbers are still allowed.
        requested_number = (request.form.get("part_number") or "").strip()
        part_number = requested_number if requested_number and not requested_number.upper().startswith("PRT-") else None
        try:
            conn.execute("BEGIN IMMEDIATE")
            if not part_number:
                part_number = next_part_number(conn)
            conn.execute("""
                INSERT INTO parts(
                    part_number,part_name,category,supplier,quantity_on_hand,reorder_level,
                    unit_cost_inc_gst,gst_amount_per_unit,storage_location,notes,
                    vehicle_id,vehicle_stock_no,vin,make,model,year,condition,selling_price,status,date_added,
                    position,fitment,manufacturer_part_no,updated_at
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP,?,?,?,CURRENT_TIMESTAMP)
            """, (
                part_number, part_name, request.form.get("category"),
                "Donor vehicle", qty, 0, cost, gst, request.form.get("storage_location"),
                request.form.get("notes"), vehicle_id, vehicle["stock_no"], vehicle["vin"],
                vehicle["make"], vehicle["model"], vehicle["year"],
                request.form.get("condition") or "Used", selling,
                request.form.get("status") or "In Stock", request.form.get("position"),
                request.form.get("fitment"), request.form.get("manufacturer_part_no"),
            ))
            conn.execute(
                "UPDATE vehicles SET vehicle_purpose='Parts Vehicle', dismantling_status=CASE WHEN COALESCE(dismantling_status,'Not Started')='Not Started' THEN 'Dismantling' ELSE dismantling_status END WHERE id=?",
                (vehicle_id,),
            )
            conn.commit()
        except sqlite3.IntegrityError as exc:
            conn.rollback()
            conn.close()
            flash(f"Could not add part: {exc}", "error")
            return redirect(url_for("vehicle_dismantling", vehicle_id=vehicle_id))
        except sqlite3.OperationalError as exc:
            conn.rollback()
            conn.close()
            flash(f"Database busy while adding part. Please try again: {exc}", "error")
            return redirect(url_for("vehicle_dismantling", vehicle_id=vehicle_id))
        conn.close()
        flash(f"{part_name} added to Parts Centre.", "success")
        return redirect(url_for("vehicle_dismantling", vehicle_id=vehicle_id))

    parts = conn.execute("SELECT * FROM parts WHERE vehicle_id=? ORDER BY id DESC", (vehicle_id,)).fetchall()
    revenue = float(conn.execute("""
        SELECT COALESCE(SUM(ps.sale_price),0) AS revenue
        FROM part_sales ps JOIN parts p ON p.id=ps.part_id
        WHERE p.vehicle_id=?
    """, (vehicle_id,)).fetchone()["revenue"] or 0)
    remaining_retail = float(conn.execute(
        "SELECT COALESCE(SUM(quantity_on_hand*selling_price),0) AS v FROM parts WHERE vehicle_id=?",
        (vehicle_id,),
    ).fetchone()["v"] or 0)
    expense_total = float(conn.execute("SELECT COALESCE(SUM(cost_inc_gst),0) AS v FROM expenses WHERE vehicle_id=?", (vehicle_id,)).fetchone()["v"] or 0)
    job_total = float(conn.execute("SELECT COALESCE(SUM(CASE WHEN actual_cost_inc_gst>0 THEN actual_cost_inc_gst ELSE estimated_cost END),0) AS v FROM job_cards WHERE vehicle_id=?", (vehicle_id,)).fetchone()["v"] or 0)
    service_total = float(conn.execute("SELECT COALESCE(SUM(cost_inc_gst),0) AS v FROM service_entries WHERE vehicle_id=?", (vehicle_id,)).fetchone()["v"] or 0)
    parts_used_total = float(conn.execute("SELECT COALESCE(SUM(quantity_used*unit_cost_inc_gst),0) AS v FROM part_usage WHERE vehicle_id=?", (vehicle_id,)).fetchone()["v"] or 0)
    invested = float(vehicle["purchase_price_inc_gst"] or 0) + expense_total + job_total + service_total + parts_used_total
    realised_revenue = revenue + float(vehicle["shell_sale_price"] or 0)
    profit = realised_revenue - invested
    recovery_pct = (realised_revenue / invested * 100) if invested > 0 else (100.0 if realised_revenue > 0 else 0.0)
    remaining_to_break_even = max(0.0, invested - realised_revenue)
    break_even_status = "Investment recovered" if realised_revenue >= invested and invested > 0 else "Recovering investment"
    conn.close()
    return render_template(
        "dismantling.html", vehicle=vehicle, parts=parts, revenue=revenue, invested=invested, profit=profit,
        remaining_retail=remaining_retail, recovery_pct=recovery_pct,
        remaining_to_break_even=remaining_to_break_even, break_even_status=break_even_status,
    )


@app.route("/parts/<int:part_id>/sell", methods=["POST"])
@login_required
def part_sell(part_id):
    conn = db()
    part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
    if not part:
        conn.close()
        flash("Part not found.", "error")
        return redirect(url_for("parts_page"))
    qty = float(request.form.get("quantity") or 1)
    if qty <= 0 or qty > float(part["quantity_on_hand"] or 0):
        conn.close()
        flash("Sale quantity is invalid.", "error")
        return redirect(request.referrer or url_for("parts_page"))
    sale_price = float(request.form.get("sale_price") or part["selling_price"] or 0)
    remaining = max(0, float(part["quantity_on_hand"] or 0) - qty)
    try:
        conn.execute("BEGIN IMMEDIATE")
        conn.execute("""
            INSERT INTO part_sales(
                part_id,quantity,customer_name,customer_phone,customer_email,sale_price,freight_cost,
                payment_method,warranty,invoice_number,notes
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?)
        """, (
            part_id, qty, request.form.get("customer_name"), request.form.get("customer_phone"),
            request.form.get("customer_email"), sale_price, float(request.form.get("freight_cost") or 0),
            request.form.get("payment_method"), request.form.get("warranty"),
            request.form.get("invoice_number"), request.form.get("notes"),
        ))
        conn.execute(
            "UPDATE parts SET quantity_on_hand=?, status=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
            (remaining, "Sold" if remaining <= 0 else "In Stock", part_id),
        )
        conn.commit()
    except (sqlite3.IntegrityError, sqlite3.OperationalError, ValueError) as exc:
        conn.rollback()
        conn.close()
        flash(f"Could not record part sale: {exc}", "error")
        return redirect(request.referrer or url_for("part_detail", part_id=part_id))
    vehicle_id = part["vehicle_id"]
    part_number = part["part_number"] or f"Part {part_id}"
    conn.close()
    log_action(
        "Part sold", "part", part_id,
        f"{part_number}; qty {qty:g}; sale ${sale_price:,.2f}; donor {part['vehicle_stock_no'] or 'none'}",
    )
    flash(f"Part sale recorded. {part_number} revenue is now included in the donor vehicle Financial Snapshot.", "success")
    if vehicle_id:
        return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))
    return redirect(url_for("part_detail", part_id=part_id))


@app.route("/parts", methods=["GET", "POST"])
@login_required
def parts_page():
    conn = db()
    if request.method == "POST":
        part_name = (request.form.get("part_name") or "").strip()
        if not part_name:
            conn.close()
            flash("Part name is required.", "error")
            return redirect(url_for("parts_page"))

        quantity = float(request.form.get("quantity_on_hand") or 0)
        unit_cost = float(request.form.get("unit_cost_inc_gst") or 0)
        gst = round(unit_cost / 11, 2)
        requested_number = (request.form.get("part_number") or "").strip()
        part_number = requested_number if requested_number and not requested_number.upper().startswith("PRT-") else None
        source_stock = (request.form.get("vehicle_stock_no") or "").strip()
        donor = None
        if source_stock:
            donor = conn.execute("SELECT * FROM vehicles WHERE stock_no=?", (source_stock,)).fetchone()

        try:
            conn.execute("BEGIN IMMEDIATE")
            if not part_number:
                part_number = next_part_number(conn)
            conn.execute("""
                INSERT INTO parts(
                    part_number,part_name,category,supplier,quantity_on_hand,reorder_level,
                    unit_cost_inc_gst,gst_amount_per_unit,storage_location,notes,
                    vehicle_id,vehicle_stock_no,vin,make,model,year,condition,selling_price,status,
                    position,fitment,manufacturer_part_no,barcode,date_added,updated_at
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)
            """, (
                part_number, part_name, request.form.get("category"), request.form.get("supplier"),
                quantity, float(request.form.get("reorder_level") or 0), unit_cost, gst,
                request.form.get("storage_location"), request.form.get("notes"),
                donor["id"] if donor else None, source_stock or None,
                donor["vin"] if donor else request.form.get("vin"),
                donor["make"] if donor else request.form.get("make"),
                donor["model"] if donor else request.form.get("model"),
                donor["year"] if donor else (request.form.get("year") or None),
                request.form.get("condition") or "Used", float(request.form.get("selling_price") or 0),
                request.form.get("status") or "In Stock", request.form.get("position"),
                request.form.get("fitment"), request.form.get("manufacturer_part_no"), (request.form.get("barcode") or part_number),
            ))
        except sqlite3.IntegrityError:
            conn.rollback()
            conn.close()
            flash(f"Part number {part_number} already exists.", "error")
            return redirect(url_for("parts_page"))
        except sqlite3.OperationalError as exc:
            conn.rollback()
            conn.close()
            flash(f"Database busy while adding part. Please try again: {exc}", "error")
            return redirect(url_for("parts_page"))
        conn.commit()
        part_id = conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
        conn.close()
        log_action("Part added", "part", part_id, f"{part_number} - {part_name}")
        flash(f"{part_number} added to Parts Centre.", "success")
        return redirect(url_for("part_detail", part_id=part_id))

    q = request.args.get("q", "").strip()
    status_filter = request.args.get("status", "").strip()
    category_filter = request.args.get("category", "").strip()
    source_filter = request.args.get("source", "").strip()

    where = []
    params = []
    if q:
        token = f"%{q}%"
        where.append("(part_number LIKE ? OR part_name LIKE ? OR category LIKE ? OR supplier LIKE ? OR storage_location LIKE ? OR vehicle_stock_no LIKE ? OR manufacturer_part_no LIKE ? OR alternate_part_numbers LIKE ? OR barcode LIKE ?)")
        params.extend([token] * 9)
    if status_filter:
        where.append("status=?")
        params.append(status_filter)
    if category_filter:
        where.append("category=?")
        params.append(category_filter)
    if source_filter:
        where.append("vehicle_stock_no=?")
        params.append(source_filter)
    sql = "SELECT * FROM parts"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY CASE WHEN status='In Stock' THEN 0 WHEN status='Reserved' THEN 1 ELSE 2 END, part_name, id DESC"
    rows = conn.execute(sql, params).fetchall()

    part_metrics = conn.execute("""
        SELECT
          COUNT(*) AS part_lines,
          COALESCE(SUM(quantity_on_hand),0) AS parts_on_hand,
          COALESCE(SUM(quantity_on_hand * unit_cost_inc_gst),0) AS stock_cost,
          COALESCE(SUM(quantity_on_hand * selling_price),0) AS retail_value,
          SUM(CASE WHEN quantity_on_hand <= reorder_level AND status NOT IN ('Sold','Scrap') THEN 1 ELSE 0 END) AS low_stock,
          SUM(CASE WHEN status='Reserved' THEN 1 ELSE 0 END) AS reserved_lines
        FROM parts
    """).fetchone()
    sales_metrics = conn.execute("""
        SELECT COALESCE(SUM(quantity),0) AS units_sold,
               COALESCE(SUM(sale_price),0) AS revenue,
               COALESCE(SUM(freight_cost),0) AS freight
        FROM part_sales
    """).fetchone()
    donor_vehicles = conn.execute("""
        SELECT v.*,
               COUNT(p.id) AS part_lines,
               COALESCE(SUM(p.quantity_on_hand),0) AS units_remaining,
               COALESCE(SUM(p.quantity_on_hand * p.selling_price),0) AS remaining_retail
        FROM vehicles v
        LEFT JOIN parts p ON p.vehicle_id=v.id
        WHERE COALESCE(v.vehicle_purpose,'Retail Sale')='Parts Vehicle' OR v.status='BER'
        GROUP BY v.id
        ORDER BY CASE WHEN COALESCE(v.dismantling_status,'Not Started')='Complete' THEN 1 ELSE 0 END, v.id DESC
    """).fetchall()
    categories = conn.execute("SELECT DISTINCT category FROM parts WHERE COALESCE(category,'')!='' ORDER BY category").fetchall()
    sources = conn.execute("SELECT DISTINCT vehicle_stock_no FROM parts WHERE COALESCE(vehicle_stock_no,'')!='' ORDER BY vehicle_stock_no").fetchall()
    next_part = next_part_number(conn)
    conn.close()
    return render_template(
        "parts.html", parts=rows, q=q, status_filter=status_filter, category_filter=category_filter,
        source_filter=source_filter, part_metrics=part_metrics, sales_metrics=sales_metrics,
        donor_vehicles=donor_vehicles, categories=categories, sources=sources, next_part=next_part,
    )


@app.route("/parts/<int:part_id>")
@login_required
def part_detail(part_id):
    conn = db()
    part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
    if not part:
        conn.close()
        return "Part not found", 404
    photos = conn.execute("SELECT * FROM part_photos WHERE part_id=? ORDER BY is_featured DESC,id DESC", (part_id,)).fetchall()
    sales = conn.execute("SELECT * FROM part_sales WHERE part_id=? ORDER BY sale_date DESC,id DESC", (part_id,)).fetchall()
    sold_qty = sum(float(r["quantity"] or 0) for r in sales)
    revenue = sum(float(r["sale_price"] or 0) for r in sales)
    freight = sum(float(r["freight_cost"] or 0) for r in sales)
    sold_cost = sold_qty * float(part["unit_cost_inc_gst"] or 0)
    gross_profit = revenue - sold_cost - freight
    donor = conn.execute("SELECT * FROM vehicles WHERE id=?", (part["vehicle_id"],)).fetchone() if part["vehicle_id"] else None
    price_suggestion = parts_price_suggestion(conn, part)
    location_history = conn.execute("SELECT * FROM part_location_history WHERE part_id=? ORDER BY moved_at DESC,id DESC LIMIT 20", (part_id,)).fetchall()
    added = None
    try:
        added = datetime.fromisoformat(str(part["date_added"] or part["updated_at"] or "").replace("Z", "+00:00"))
    except (ValueError, TypeError):
        added = None
    days_in_stock = max(0, (datetime.now().date() - added.date()).days) if added else 0
    fields = [part["part_number"], part["part_name"], part["category"], part["manufacturer_part_no"], part["storage_location"], part["fitment"], part["condition"], part["barcode"]]
    profile_percent = round(sum(1 for value in fields if str(value or "").strip()) / len(fields) * 100)
    current_price = float(part["selling_price"] or 0)
    unit_cost = float(part["unit_cost_inc_gst"] or 0)
    margin = current_price - unit_cost
    margin_percent = round((margin / current_price * 100), 1) if current_price else 0
    intelligence = {"days_in_stock": days_in_stock, "profile_percent": profile_percent, "margin": margin, "margin_percent": margin_percent}
    conn.close()
    return render_template("part_detail.html", part=part, photos=photos, sales=sales, donor=donor,
                           sold_qty=sold_qty, revenue=revenue, freight=freight, gross_profit=gross_profit, price_suggestion=price_suggestion,
                           location_history=location_history, intelligence=intelligence)


@app.route("/parts/<int:part_id>/edit", methods=["GET", "POST"])
@login_required
def part_edit(part_id):
    conn = db()
    part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
    if not part:
        conn.close()
        return "Part not found", 404
    if request.method == "POST":
        part_number = (request.form.get("part_number") or "").strip() or part["part_number"] or next_part_number(conn)
        source_stock = (request.form.get("vehicle_stock_no") or "").strip()
        donor = conn.execute("SELECT * FROM vehicles WHERE stock_no=?", (source_stock,)).fetchone() if source_stock else None
        try:
            conn.execute("""
                UPDATE parts SET
                    part_number=?,part_name=?,category=?,subcategory=?,description=?,supplier=?,quantity_on_hand=?,reorder_level=?,
                    unit_cost_inc_gst=?,gst_amount_per_unit=?,selling_price=?,storage_location=?,notes=?,vehicle_id=?,vehicle_stock_no=?,
                    vin=?,make=?,model=?,year=?,condition=?,status=?,engine_code=?,transmission_code=?,barcode=?,position=?,fitment=?,
                    manufacturer_part_no=?,reserved_for=?,reserved_until=?,alternate_part_numbers=?,interchange_notes=?,warranty_days=?,
                    weight_kg=?,length_cm=?,width_cm=?,height_cm=?,inventory_last_checked=?,updated_at=CURRENT_TIMESTAMP
                WHERE id=?
            """, (
                part_number, request.form.get("part_name"), request.form.get("category"), request.form.get("subcategory"),
                request.form.get("description"), request.form.get("supplier"), float(request.form.get("quantity_on_hand") or 0),
                float(request.form.get("reorder_level") or 0), float(request.form.get("unit_cost_inc_gst") or 0),
                round(float(request.form.get("unit_cost_inc_gst") or 0)/11, 2), float(request.form.get("selling_price") or 0),
                request.form.get("storage_location"), request.form.get("notes"), donor["id"] if donor else part["vehicle_id"],
                source_stock or None, donor["vin"] if donor else request.form.get("vin"), donor["make"] if donor else request.form.get("make"),
                donor["model"] if donor else request.form.get("model"), donor["year"] if donor else (request.form.get("year") or None),
                request.form.get("condition") or "Used", request.form.get("status") or "In Stock", request.form.get("engine_code"),
                request.form.get("transmission_code"), (request.form.get("barcode") or part_number), request.form.get("position"), request.form.get("fitment"),
                request.form.get("manufacturer_part_no"), request.form.get("reserved_for"), request.form.get("reserved_until") or None,
                request.form.get("alternate_part_numbers"), request.form.get("interchange_notes"), int(float(request.form.get("warranty_days") or 0)),
                float(request.form.get("weight_kg") or 0), float(request.form.get("length_cm") or 0), float(request.form.get("width_cm") or 0),
                float(request.form.get("height_cm") or 0), request.form.get("inventory_last_checked") or None, part_id,
            ))
        except sqlite3.IntegrityError:
            conn.rollback()
            conn.close()
            flash(f"Part number {part_number} already exists.", "error")
            return redirect(url_for("part_edit", part_id=part_id))
        except sqlite3.OperationalError as exc:
            conn.rollback()
            conn.close()
            flash(f"Could not update part: {exc}", "error")
            return redirect(url_for("part_edit", part_id=part_id))
        conn.commit()
        conn.close()
        log_action(
            "Part updated", "part", part_id,
            f"{part_number}; {request.form.get('part_name') or ''}; qty {request.form.get('quantity_on_hand') or 0}; status {request.form.get('status') or 'In Stock'}",
        )
        flash("Part updated.", "success")
        return redirect(url_for("part_detail", part_id=part_id))
    donor_vehicles = conn.execute("SELECT id,stock_no,year,make,model FROM vehicles WHERE vehicle_purpose='Parts Vehicle' OR status='BER' ORDER BY stock_no").fetchall()
    conn.close()
    return render_template("part_edit.html", part=part, donor_vehicles=donor_vehicles)


@app.route("/parts/<int:part_id>/photos", methods=["POST"])
@login_required
def part_photo_add(part_id):
    conn = db()
    part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
    if not part:
        conn.close()
        return "Part not found", 404
    files = request.files.getlist("photos")
    added = 0
    image_exts = {"png", "jpg", "jpeg", "webp"}
    for file in files:
        if not file or not file.filename:
            continue
        ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""
        if ext not in image_exts:
            continue
        filename = save_upload(file)
        featured = 1 if added == 0 and not conn.execute("SELECT id FROM part_photos WHERE part_id=? LIMIT 1", (part_id,)).fetchone() else 0
        conn.execute("INSERT INTO part_photos(part_id,filename,caption,is_featured) VALUES(?,?,?,?)",
                     (part_id, filename, request.form.get("caption"), featured))
        added += 1
    conn.commit()
    conn.close()
    flash(f"{added} part photo(s) uploaded." if added else "No valid images selected.", "success" if added else "error")
    return redirect(url_for("part_detail", part_id=part_id))


@app.route("/parts/<int:part_id>/photos/<int:photo_id>/feature", methods=["POST"])
@login_required
def part_photo_feature(part_id, photo_id):
    conn = db()
    conn.execute("UPDATE part_photos SET is_featured=0 WHERE part_id=?", (part_id,))
    conn.execute("UPDATE part_photos SET is_featured=1 WHERE id=? AND part_id=?", (photo_id, part_id))
    conn.commit(); conn.close()
    flash("Featured part photo updated.", "success")
    return redirect(url_for("part_detail", part_id=part_id))


@app.route("/parts/<int:part_id>/photos/<int:photo_id>/delete", methods=["POST"])
@login_required
def part_photo_delete(part_id, photo_id):
    conn = db()
    photo = conn.execute("SELECT * FROM part_photos WHERE id=? AND part_id=?", (photo_id, part_id)).fetchone()
    if photo:
        conn.execute("DELETE FROM part_photos WHERE id=?", (photo_id,))
        conn.commit()
    conn.close()
    if photo:
        try:
            (UPLOAD_DIR / photo["filename"]).unlink(missing_ok=True)
        except OSError:
            pass
    flash("Part photo deleted.", "success")
    return redirect(url_for("part_detail", part_id=part_id))


@app.route("/vehicles/<int:vehicle_id>/use-part", methods=["POST"])
@login_required
def use_part(vehicle_id):
    part_id = int(request.form.get("part_id"))
    quantity = float(request.form.get("quantity_used") or 0)
    conn = db()
    part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()

    if not part:
        conn.close()
        flash("Part not found.", "error")
        return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))

    if quantity <= 0:
        conn.close()
        flash("Quantity must be greater than zero.", "error")
        return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))

    if quantity > part["quantity_on_hand"]:
        conn.close()
        flash("Not enough stock available.", "error")
        return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))

    conn.execute("""
        INSERT INTO part_usage(
            part_id,vehicle_id,job_card_id,usage_date,quantity_used,unit_cost_inc_gst,paid_by,notes
        ) VALUES(?,?,?,?,?,?,?,?)
    """, (
        part_id,
        vehicle_id,
        int(request.form.get("job_card_id")) if request.form.get("job_card_id") else None,
        request.form.get("usage_date"),
        quantity,
        part["unit_cost_inc_gst"],
        request.form.get("paid_by"),
        request.form.get("notes"),
    ))
    conn.execute(
        "UPDATE parts SET quantity_on_hand=quantity_on_hand-? WHERE id=?",
        (quantity, part_id),
    )
    conn.commit()
    conn.close()
    flash("Part allocated to vehicle.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/timeline")
@login_required
def vehicle_timeline(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close()
        return "Vehicle not found", 404

    events = []
    def add_event(event_date, event_type, title, details="", amount=None, status=""):
        if event_date:
            events.append({
                "date": str(event_date), "type": event_type, "title": title,
                "details": details or "", "amount": amount, "status": status or ""
            })

    add_event(vehicle["created_at"], "Vehicle", "Vehicle created", f"{vehicle['stock_no']} - {vehicle['year'] or ''} {vehicle['make']} {vehicle['model']}")
    add_event(vehicle["purchase_date"], "Purchase", "Vehicle purchased", "Purchase recorded", vehicle["purchase_price_inc_gst"])

    for row in conn.execute("SELECT * FROM expenses WHERE vehicle_id=?", (vehicle_id,)).fetchall():
        add_event(row["expense_date"], "Expense", row["category"] or "Expense", row["description"], row["cost_inc_gst"])
    for row in conn.execute("SELECT * FROM job_cards WHERE vehicle_id=?", (vehicle_id,)).fetchall():
        amount = row["actual_cost_inc_gst"] if float(row["actual_cost_inc_gst"] or 0) > 0 else row["estimated_cost"]
        add_event(row["job_date"], "Workshop", row["description"], row["category"] or "Workshop job", amount, row["status"] or "")
    for row in conn.execute("SELECT * FROM service_entries WHERE vehicle_id=?", (vehicle_id,)).fetchall():
        add_event(row["service_date"], "Service", row["service_type"] or "Service", row["description"], row["cost_inc_gst"])
    for row in conn.execute("SELECT * FROM reminders WHERE vehicle_id=?", (vehicle_id,)).fetchall():
        add_event(row["reminder_date"], "Reminder", row["title"], row["reminder_type"], None, "Completed" if row["completed"] else "Open")
    for row in conn.execute("SELECT * FROM tasks WHERE vehicle_id=?", (vehicle_id,)).fetchall():
        add_event(row["task_date"] or row["created_at"], "Task", row["title"], row["category"] or "", None, row["status"] or "")
    for row in conn.execute("SELECT * FROM parts WHERE vehicle_id=?", (vehicle_id,)).fetchall():
        add_event(row["date_added"] or row["updated_at"], "Dismantling", f"Part added: {row['part_number'] or ''} {row['part_name']}", row["storage_location"] or "", row["selling_price"], row["status"] or "")
    for row in conn.execute("""
        SELECT ps.*,p.part_number,p.part_name FROM part_sales ps
        JOIN parts p ON p.id=ps.part_id WHERE p.vehicle_id=?
    """, (vehicle_id,)).fetchall():
        add_event(row["sale_date"], "Part Sale", f"Sold {row['part_number'] or ''} {row['part_name']}", row["customer_name"] or "", row["sale_price"], "Sold")
    sale = conn.execute("SELECT * FROM sales WHERE vehicle_id=?", (vehicle_id,)).fetchone()
    if sale:
        add_event(sale["sale_date"], "Vehicle Sale", f"Vehicle sold to {sale['buyer_name'] or 'buyer'}", sale["invoice_number"] or "", sale["sale_price_inc_gst"], sale["invoice_status"] or "")
    for row in conn.execute("SELECT * FROM audit_log WHERE entity_type='vehicle' AND entity_id=?", (vehicle_id,)).fetchall():
        add_event(row["created_at"], "Activity", row["action"], row["details"] or "")

    events.sort(key=lambda x: x["date"] or "", reverse=True)
    conn.close()
    return render_template("vehicle_timeline.html", vehicle=vehicle, events=events)


@app.route("/search")
@login_required
def global_search():
    q = request.args.get("q", "").strip()
    conn = db()
    vehicles = []
    contacts = []
    parts = []
    invoices = []
    if q:
        pattern = f"%{q}%"
        vehicles = conn.execute("""
            SELECT * FROM vehicles
            WHERE stock_no LIKE ? OR vin LIKE ? OR registration LIKE ? OR make LIKE ? OR model LIKE ?
            ORDER BY id DESC LIMIT 50
        """, (pattern, pattern, pattern, pattern, pattern)).fetchall()
        contacts = conn.execute("""
            SELECT * FROM contacts
            WHERE name LIKE ? OR phone LIKE ? OR email LIKE ?
            ORDER BY name LIMIT 50
        """, (pattern, pattern, pattern)).fetchall()
        parts = conn.execute("""
            SELECT * FROM parts
            WHERE part_number LIKE ? OR part_name LIKE ? OR manufacturer_part_no LIKE ?
               OR barcode LIKE ? OR vehicle_stock_no LIKE ? OR storage_location LIKE ? OR fitment LIKE ?
            ORDER BY id DESC LIMIT 50
        """, (pattern, pattern, pattern, pattern, pattern, pattern, pattern)).fetchall()
        invoices = conn.execute("""
            SELECT s.*,v.stock_no,v.make,v.model,v.year FROM sales s
            JOIN vehicles v ON v.id=s.vehicle_id
            WHERE s.invoice_number LIKE ? OR s.buyer_name LIKE ? OR s.buyer_phone LIKE ?
               OR v.stock_no LIKE ? OR v.registration LIKE ?
            ORDER BY s.id DESC LIMIT 50
        """, (pattern, pattern, pattern, pattern, pattern)).fetchall()
    conn.close()
    return render_template("global_search.html", q=q, vehicles=vehicles, contacts=contacts, parts=parts, invoices=invoices)

@app.route("/facebook-marketplace")
@login_required
def facebook_marketplace():
    return redirect("https://www.facebook.com/marketplace/")


@app.route("/reports/finance")
@login_required
def finance_summary():
    conn = db()
    rows = conn.execute("""
        SELECT v.id,v.stock_no,v.year,v.make,v.model,v.status,
               v.purchase_price_inc_gst,v.barry_contribution,v.matt_contribution,
               COALESCE((SELECT SUM(e.cost_inc_gst) FROM expenses e WHERE e.vehicle_id=v.id),0) AS expenses,
               COALESCE((SELECT SUM(CASE WHEN j.actual_cost_inc_gst>0 THEN j.actual_cost_inc_gst ELSE j.estimated_cost END)
                         FROM job_cards j WHERE j.vehicle_id=v.id),0) AS jobs,
               COALESCE((SELECT SUM(s.cost_inc_gst) FROM service_entries s WHERE s.vehicle_id=v.id),0) AS services,
               COALESCE((SELECT SUM(u.quantity_used*u.unit_cost_inc_gst) FROM part_usage u WHERE u.vehicle_id=v.id),0) AS parts,
               COALESCE((SELECT sale_price_inc_gst FROM sales s WHERE s.vehicle_id=v.id),0) AS sale_price
        FROM vehicles v
        ORDER BY v.id DESC
    """).fetchall()
    conn.close()

    report = []
    total_barry = 0
    total_matt = 0
    total_profit = 0
    money_in_stock = 0

    for r in rows:
        total_cost = r["purchase_price_inc_gst"] + r["expenses"] + r["jobs"] + r["services"] + r["parts"]
        profit = r["sale_price"] - total_cost
        if r["status"] not in ("Sold", "BER"):
            money_in_stock += total_cost
        total_barry += r["barry_contribution"]
        total_matt += r["matt_contribution"]
        total_profit += profit
        report.append(dict(r) | {"total_cost": total_cost, "profit": profit})

    return render_template(
        "finance_summary.html",
        rows=report,
        total_barry=total_barry,
        total_matt=total_matt,
        total_profit=total_profit,
        money_in_stock=money_in_stock,
        barry_profit_share=total_profit/2,
        matt_profit_share=total_profit/2,
    )


@app.route("/reports/performance")
@login_required
def performance_report():
    conn = db()
    monthly = conn.execute("""
        SELECT substr(sale_date,1,7) AS month,
               COUNT(*) AS vehicles_sold,
               SUM(sale_price_inc_gst) AS sales_total
        FROM sales
        WHERE sale_date IS NOT NULL AND sale_date!=''
        GROUP BY substr(sale_date,1,7)
        ORDER BY month
    """).fetchall()

    makes = conn.execute("""
        SELECT make, COUNT(*) AS total
        FROM vehicles
        GROUP BY make
        ORDER BY total DESC, make
        LIMIT 10
    """).fetchall()
    conn.close()

    max_sales = max([r["sales_total"] or 0 for r in monthly], default=1)
    return render_template("performance_report.html", monthly=monthly, makes=makes, max_sales=max_sales)


@app.route("/vehicles/<int:vehicle_id>/contract")
@login_required
def sale_contract(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    sale = conn.execute("SELECT * FROM sales WHERE vehicle_id=?", (vehicle_id,)).fetchone()
    conn.close()
    if not vehicle or not sale:
        return "Sales contract is not available until a sale is recorded.", 404
    balance_due = sale["sale_price_inc_gst"] - sale["deposit_amount"] - sale["trade_in_value"]
    return render_template(
        "sale_contract.html",
        vehicle=vehicle,
        sale=sale,
        balance_due=balance_due,
    )


@app.route("/customers")
@login_required
def customer_history():
    conn = db()
    q = request.args.get("q", "").strip()
    pattern = f"%{q}%"
    vehicle_sql = """
        SELECT 'Vehicle' AS purchase_type,s.id AS sale_id,s.vehicle_id,s.sale_date AS purchase_date,
               s.buyer_name AS customer_name,s.buyer_phone AS customer_phone,s.buyer_email AS customer_email,
               s.sale_price_inc_gst AS amount,s.invoice_number,
               v.stock_no AS item_no,(COALESCE(v.year,'') || ' ' || v.make || ' ' || v.model) AS item_name,
               v.registration
        FROM sales s JOIN vehicles v ON v.id=s.vehicle_id
    """
    part_sql = """
        SELECT 'Part' AS purchase_type,ps.id AS sale_id,p.vehicle_id,ps.sale_date AS purchase_date,
               ps.customer_name,ps.customer_phone,ps.customer_email,ps.sale_price AS amount,ps.invoice_number,
               p.part_number AS item_no,p.part_name AS item_name,'' AS registration
        FROM part_sales ps JOIN parts p ON p.id=ps.part_id
    """
    params = []
    if q:
        vehicle_sql += " WHERE s.buyer_name LIKE ? OR s.buyer_phone LIKE ? OR s.buyer_email LIKE ?"
        part_sql += " WHERE ps.customer_name LIKE ? OR ps.customer_phone LIKE ? OR ps.customer_email LIKE ?"
        params = [pattern, pattern, pattern]
    vehicle_rows = conn.execute(vehicle_sql + " ORDER BY purchase_date DESC", params).fetchall()
    part_rows = conn.execute(part_sql + " ORDER BY purchase_date DESC", params).fetchall()
    conn.close()

    transactions = [dict(r) for r in vehicle_rows] + [dict(r) for r in part_rows]
    transactions.sort(key=lambda r: str(r.get("purchase_date") or ""), reverse=True)
    customers = {}
    for row in transactions:
        email = (row.get("customer_email") or "").strip().lower()
        phone = re.sub(r"\s+", "", row.get("customer_phone") or "")
        name = (row.get("customer_name") or "Unknown customer").strip()
        key = email or phone or name.lower()
        profile = customers.setdefault(key, {"name": name, "phone": row.get("customer_phone") or "", "email": row.get("customer_email") or "", "total_spend": 0.0, "vehicle_count": 0, "part_count": 0, "last_purchase": "", "transactions": []})
        profile["total_spend"] += float(row.get("amount") or 0)
        profile["vehicle_count"] += 1 if row["purchase_type"] == "Vehicle" else 0
        profile["part_count"] += 1 if row["purchase_type"] == "Part" else 0
        profile["last_purchase"] = max(profile["last_purchase"], str(row.get("purchase_date") or ""))
        profile["transactions"].append(row)
    customer_rows = sorted(customers.values(), key=lambda x: (x["total_spend"], x["last_purchase"]), reverse=True)
    return render_template("customer_history.html", customers=customer_rows, transactions=transactions, q=q)


@app.route("/reports/bas")
@login_required
def bas_report():
    start = request.args.get("start", "")
    end = request.args.get("end", "")
    conn = db()

    clauses = []
    params = []
    if start:
        clauses.append("date_value>=?")
        params.append(start)
    if end:
        clauses.append("date_value<=?")
        params.append(end)
    where = "WHERE " + " AND ".join(clauses) if clauses else ""

    sales_rows = conn.execute(f"""
        SELECT sale_date AS date_value,sale_price_inc_gst,sale_gst
        FROM sales
        {where.replace('date_value','sale_date')}
    """, params).fetchall()

    vehicle_rows = conn.execute(f"""
        SELECT purchase_date AS date_value,purchase_price_inc_gst,purchase_gst
        FROM vehicles
        {where.replace('date_value','purchase_date')}
    """, params).fetchall()

    expense_rows = conn.execute(f"""
        SELECT expense_date AS date_value,cost_inc_gst,gst_amount
        FROM expenses
        {where.replace('date_value','expense_date')}
    """, params).fetchall()

    job_rows = conn.execute(f"""
        SELECT job_date AS date_value,
               CASE WHEN actual_cost_inc_gst>0 THEN actual_cost_inc_gst ELSE estimated_cost END AS cost_value,
               gst_amount
        FROM job_cards
        {where.replace('date_value','job_date')}
    """, params).fetchall()

    service_rows = conn.execute(f"""
        SELECT service_date AS date_value,cost_inc_gst,gst_amount
        FROM service_entries
        {where.replace('date_value','service_date')}
    """, params).fetchall()

    part_rows = conn.execute(f"""
        SELECT usage_date AS date_value,quantity_used*unit_cost_inc_gst AS cost_value,
               (quantity_used*unit_cost_inc_gst)/11.0 AS gst_amount
        FROM part_usage
        {where.replace('date_value','usage_date')}
    """, params).fetchall()
    conn.close()

    sales_total = sum(r["sale_price_inc_gst"] for r in sales_rows)
    gst_collected = sum(r["sale_gst"] for r in sales_rows)
    purchase_total = sum(r["purchase_price_inc_gst"] for r in vehicle_rows)
    gst_vehicle = sum(r["purchase_gst"] for r in vehicle_rows)
    expense_total = sum(r["cost_inc_gst"] for r in expense_rows)
    gst_expenses = sum(r["gst_amount"] for r in expense_rows)
    job_total = sum(r["cost_value"] for r in job_rows)
    gst_jobs = sum(r["gst_amount"] for r in job_rows)
    service_total = sum(r["cost_inc_gst"] for r in service_rows)
    gst_services = sum(r["gst_amount"] for r in service_rows)
    parts_total = sum(r["cost_value"] for r in part_rows)
    gst_parts = sum(r["gst_amount"] for r in part_rows)

    gst_paid = gst_vehicle + gst_expenses + gst_jobs + gst_services + gst_parts
    net_gst = gst_collected - gst_paid

    return render_template(
        "bas_report.html",
        start=start,
        end=end,
        sales_total=sales_total,
        gst_collected=gst_collected,
        purchase_total=purchase_total,
        expense_total=expense_total,
        job_total=job_total,
        service_total=service_total,
        parts_total=parts_total,
        gst_paid=gst_paid,
        net_gst=net_gst,
    )


@app.route("/audit")
@login_required
def audit_page():
    conn = db()
    rows = conn.execute("""
        SELECT * FROM audit_log
        ORDER BY created_at DESC,id DESC
        LIMIT 500
    """).fetchall()
    conn.close()
    return render_template("audit_log.html", rows=rows)


@app.route("/quick-add-vehicle", methods=["POST"])
@login_required
def quick_add_vehicle():
    conn = db()
    stock_no = next_stock_number(conn)
    price = float(request.form.get("purchase_price_inc_gst") or 0)
    gst = round(price / 11, 2)

    cursor = conn.execute("""
        INSERT INTO vehicles(
            stock_no,status,purchase_date,make,model,variant,year,vin,registration,
            odometer_km,colour,purchase_price_inc_gst,purchase_gst,
            barry_contribution,matt_contribution,sale_ownership,rego_expiry,photo_filename,notes,
            ppsr_number,roadworthy_status,service_due_date,service_history
        ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        stock_no,
        request.form.get("status") or "In Stock",
        request.form.get("purchase_date"),
        request.form.get("make"),
        request.form.get("model"),
        request.form.get("variant"),
        int(request.form.get("year") or 0) or None,
        request.form.get("vin") or None,
        request.form.get("registration"),
        int(request.form.get("odometer_km") or 0) or None,
        request.form.get("colour"),
        price,
        gst,
        float(request.form.get("barry_contribution") or 0),
        float(request.form.get("matt_contribution") or 0),
        request.form.get("sale_ownership") or "BAM Joint",
        request.form.get("rego_expiry"),
        None,
        request.form.get("notes"),
        request.form.get("ppsr_number"),
        request.form.get("roadworthy_status") or "Not Checked",
        request.form.get("service_due_date"),
        request.form.get("service_history"),
    ))
    vehicle_id = cursor.lastrowid
    conn.commit()
    conn.close()

    log_action("Quick vehicle added", "vehicle", vehicle_id, stock_no)
    flash(f"{stock_no} added successfully.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/photos/featured/<int:photo_id>", methods=["POST"])
@login_required
def set_featured_photo(vehicle_id, photo_id):
    conn = db()
    photo = conn.execute(
        "SELECT * FROM vehicle_photos WHERE id=? AND vehicle_id=?",
        (photo_id, vehicle_id),
    ).fetchone()
    if not photo:
        conn.close()
        flash("Photo not found.", "error")
        return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))

    conn.execute(
        "UPDATE vehicles SET featured_photo_id=?,photo_filename=? WHERE id=?",
        (photo_id, photo["filename"], vehicle_id),
    )
    conn.commit()
    conn.close()
    log_action("Featured photo changed", "vehicle", vehicle_id, photo["filename"])
    flash("Featured photo updated.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/photos/delete/<int:photo_id>", methods=["POST"])
@login_required
def delete_vehicle_photo(vehicle_id, photo_id):
    conn = db()
    photo = conn.execute(
        "SELECT * FROM vehicle_photos WHERE id=? AND vehicle_id=?",
        (photo_id, vehicle_id),
    ).fetchone()
    if not photo:
        conn.close()
        flash("Photo not found.", "error")
        return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))

    conn.execute("DELETE FROM vehicle_photos WHERE id=?", (photo_id,))
    vehicle = conn.execute("SELECT featured_photo_id FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if vehicle and vehicle["featured_photo_id"] == photo_id:
        replacement = conn.execute(
            "SELECT id,filename FROM vehicle_photos WHERE vehicle_id=? ORDER BY id DESC LIMIT 1",
            (vehicle_id,),
        ).fetchone()
        conn.execute(
            "UPDATE vehicles SET featured_photo_id=?,photo_filename=? WHERE id=?",
            (replacement["id"] if replacement else None, replacement["filename"] if replacement else None, vehicle_id),
        )
    conn.commit()
    conn.close()

    try:
        file_path = UPLOAD_DIR / photo["filename"]
        if file_path.exists():
            file_path.unlink()
    except OSError:
        pass

    log_action("Vehicle photo deleted", "vehicle", vehicle_id, photo["filename"])
    flash("Photo deleted.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.post("/vehicles/<int:vehicle_id>/automatic-valuation")
@login_required
def vehicle_automatic_valuation(vehicle_id):
    conn=db(); vehicle=conn.execute("SELECT * FROM vehicles WHERE id=?",(vehicle_id,)).fetchone()
    if not vehicle: conn.close(); return "Vehicle not found",404
    try:
        result=bam_live_market_valuation(dict(vehicle)); prices=(result["prices"]+[0,0,0,0,0])[:5]
        source=result["summary"] + " Asking prices are market evidence, not confirmed sale prices."
        conn.execute("""UPDATE vehicles SET comparable_price_1=?,comparable_price_2=?,comparable_price_3=?,comparable_price_4=?,comparable_price_5=?,market_price_low=?,market_price_mid=?,market_price_high=?,private_value_low=?,private_value_high=?,wholesale_value_low=?,wholesale_value_high=?,trade_value_low=?,trade_value_high=?,dealer_value_low=?,dealer_value_high=?,estimated_sale_price=?,minimum_sale_price=?,valuation_provider=?,valuation_confidence=?,valuation_source=?,market_price_checked_at=CURRENT_TIMESTAMP WHERE id=?""",(*prices,result["market_low"],result["market_mid"],result["market_high"],result["private_low"],result["private_high"],result["wholesale_low"],result["wholesale_high"],result["trade_low"],result["trade_high"],result["dealer_low"],result["dealer_high"],result["market_mid"],result["wholesale_high"],"BAM AI live Australian market search",result["confidence"],source,vehicle_id))
        conn.commit(); flash(f"BAM Automatic Market Valuation saved from {len(result['prices'])} current Australian comparable listing(s).","success")
    except Exception as exc:
        flash(f"Automatic market valuation could not complete: {exc}","error")
    finally: conn.close()
    return redirect(url_for("vehicle_valuation",vehicle_id=vehicle_id))

@app.route("/vehicles/<int:vehicle_id>/valuation", methods=["GET", "POST"])
@login_required
def vehicle_valuation(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close(); return "Vehicle not found", 404
    expenses = conn.execute("SELECT COALESCE(SUM(cost_inc_gst),0) AS v FROM expenses WHERE vehicle_id=?", (vehicle_id,)).fetchone()["v"]
    jobs = conn.execute("SELECT COALESCE(SUM(CASE WHEN actual_cost_inc_gst>0 THEN actual_cost_inc_gst ELSE estimated_cost END),0) AS v FROM job_cards WHERE vehicle_id=?", (vehicle_id,)).fetchone()["v"]
    services = conn.execute("SELECT COALESCE(SUM(cost_inc_gst),0) AS v FROM service_entries WHERE vehicle_id=?", (vehicle_id,)).fetchone()["v"]
    parts = conn.execute("SELECT COALESCE(SUM(quantity_used*unit_cost_inc_gst),0) AS v FROM part_usage WHERE vehicle_id=?", (vehicle_id,)).fetchone()["v"]
    if request.method == "POST":
        prices=[]; raw_prices=[]
        for i in range(1,6):
            try: value=float((request.form.get(f"comparable_price_{i}") or "0").strip())
            except ValueError: value=0
            raw_prices.append(value)
            if value>0: prices.append(value)
        if prices:
            ordered=sorted(prices); n=len(ordered); low=ordered[0]; high=ordered[-1]
            mid=ordered[n//2] if n%2 else (ordered[n//2-1]+ordered[n//2])/2
            confidence="Good" if n>=4 else ("Limited" if n>=2 else "Single comparable")
            conn.execute("""UPDATE vehicles SET comparable_price_1=?,comparable_price_2=?,comparable_price_3=?,comparable_price_4=?,comparable_price_5=?,market_price_low=?,market_price_mid=?,market_price_high=?,private_value_low=?,private_value_high=?,wholesale_value_low=?,wholesale_value_high=?,trade_value_low=?,trade_value_high=?,dealer_value_low=?,dealer_value_high=?,estimated_sale_price=?,minimum_sale_price=?,valuation_provider=?,valuation_confidence=?,valuation_source=?,market_price_checked_at=CURRENT_TIMESTAMP WHERE id=?""", (*raw_prices,low,mid,high,low*.92,high*.97,low*.62,mid*.72,low*.68,mid*.78,low,high,mid,mid*.72,"BAM comparable market analysis",confidence,f"{n} advertised Australian comparable price(s); asking prices are not confirmed sales.",vehicle_id))
            conn.commit(); flash("Market valuation calculated. Market Mid has been put into Quick-Sale / Estimated Sale Value.","success")
        else: flash("Enter at least one comparable market price first.","error")
        vehicle=conn.execute("SELECT * FROM vehicles WHERE id=?",(vehicle_id,)).fetchone()
    conn.close()
    total_cost=float(vehicle["purchase_price_inc_gst"] or 0)+float(expenses or 0)+float(jobs or 0)+float(services or 0)+float(parts or 0)
    asset=(vehicle["asset_type"] or "Car").strip(); terms=[vehicle["year"],vehicle["make"],vehicle["model"],vehicle["variant"]]
    if asset.lower()=="boat": terms += [f'{vehicle["length_m"]}m' if vehicle["length_m"] else None,vehicle["engine_make"],f'{vehicle["horsepower"]}hp' if vehicle["horsepower"] else None,f'{vehicle["engine_hours"]} hours' if vehicle["engine_hours"] else None,"Boatsonline Gumtree Facebook Marketplace"]
    elif asset.lower()=="caravan": terms += [f'{vehicle["length_m"]}m' if vehicle["length_m"] else None,f'{vehicle["berths"]} berth' if vehicle["berths"] else None,f'{vehicle["tare_weight_kg"]}kg tare' if vehicle["tare_weight_kg"] else None,"Caravancampingsales Gumtree Facebook Marketplace"]
    elif asset.lower()=="trailer": terms += [f'{vehicle["length_m"]}m' if vehicle["length_m"] else None,f'{vehicle["tare_weight_kg"]}kg tare' if vehicle["tare_weight_kg"] else None,f'{vehicle["atm_kg"]}kg ATM' if vehicle["atm_kg"] else None,"Gumtree Facebook Marketplace"]
    else: terms += [f'{vehicle["odometer_km"]} km' if vehicle["odometer_km"] else None,"Carsales Gumtree Facebook Marketplace"]
    google_url="https://www.google.com/search?q="+urllib.parse.quote_plus(" ".join(str(x) for x in terms if x)+f" {asset} for sale Australia price")
    purchase=float(vehicle["purchase_price_inc_gst"] or 0); running_costs=max(0,total_cost-purchase); quick=float(vehicle["estimated_sale_price"] or vehicle["market_price_mid"] or 0); target_profit=float(vehicle["target_profit"] or 0)
    target_buy=max(0,float(vehicle["wholesale_value_low"] or quick*.62)-running_costs-target_profit); max_buy=max(0,float(vehicle["wholesale_value_high"] or quick*.72)-running_costs-target_profit); walk_away=max(0,quick-running_costs-target_profit); expected_profit=quick-total_cost if quick else 0
    if not quick: deal_score="-"; deal_message="Enter comparable prices to calculate the BAM buying guide."
    elif purchase<=0: deal_score="READY"; deal_message="Valuation ready. Purchase price has not been recorded yet."
    elif purchase<=target_buy: deal_score="STRONG BUY"; deal_message="Purchase price was at or below BAM Target Buy."
    elif purchase<=max_buy: deal_score="GOOD BUY"; deal_message="Purchase price was within BAM recommended buying range."
    elif purchase<=walk_away: deal_score="CAUTION"; deal_message="Purchase price was above BAM recommended buy range but below Walk-Away."
    else: deal_score="ABOVE WALK-AWAY"; deal_message="Recorded purchase price is above the calculated Walk-Away Price."
    template='''{% extends "base.html" %}{% block content %}
<div class="panel"><h1>🇦🇺 BAM Inventory Market Valuation</h1><p><b>{{vehicle.stock_no}}</b> — {{vehicle.year or ''}} {{vehicle.make}} {{vehicle.model}} · {{vehicle.asset_type or 'Car'}}</p><div class="actions"><form method="post" action="{{url_for('vehicle_automatic_valuation',vehicle_id=vehicle.id)}}" style="display:inline"><button class="btn good">✨ BAM Automatic Market Valuation</button></form><a class="btn secondary" target="_blank" rel="noopener" href="{{google_url}}">🔎 Open Google Cross-Check ↗</a><form method="post" action="{{url_for('vehicle_market_suggestion_save',vehicle_id=vehicle.id)}}" style="display:inline"><button class="btn secondary">BAM Internal Market Suggestion</button></form></div><div class="muted">BAM Automatic Market Valuation researches current Australian advertised listings, saves up to five comparable prices, fills all valuation ranges and Quick-Sale Value, then updates the buying guide. Google remains a manual cross-check.</div></div>
<form method="post"><div class="panel"><h2>📊 Comparable Market Prices</h2><div class="grid">{% for i in range(1,6) %}<div><label>Comparable {{i}} ($)</label><input type="number" step=".01" min="0" name="comparable_price_{{i}}" value="{{ vehicle['comparable_price_' ~ i] or '' }}"></div>{% endfor %}</div><div class="actions"><button class="btn good">Calculate Market Valuation & Put Into Quick-Sale Value</button></div></div></form>
<div class="panel"><h2>Market Valuation</h2><div class="valuation-grid"><div class="valuation-card"><span>Market Low</span><strong>${{'{:,.0f}'.format(vehicle.market_price_low or 0)}}</strong></div><div class="valuation-card"><span>Market Mid / Quick-Sale</span><strong>${{'{:,.0f}'.format(vehicle.market_price_mid or 0)}}</strong></div><div class="valuation-card"><span>Market High</span><strong>${{'{:,.0f}'.format(vehicle.market_price_high or 0)}}</strong></div><div class="valuation-card"><span>Wholesale</span><strong>${{'{:,.0f}'.format(vehicle.wholesale_value_low or 0)}} – ${{'{:,.0f}'.format(vehicle.wholesale_value_high or 0)}}</strong></div><div class="valuation-card"><span>Trade</span><strong>${{'{:,.0f}'.format(vehicle.trade_value_low or 0)}} – ${{'{:,.0f}'.format(vehicle.trade_value_high or 0)}}</strong></div><div class="valuation-card"><span>Private</span><strong>${{'{:,.0f}'.format(vehicle.private_value_low or 0)}} – ${{'{:,.0f}'.format(vehicle.private_value_high or 0)}}</strong></div></div><p><b>Quick-Sale / Estimated Sale Value:</b> ${{'{:,.0f}'.format(vehicle.estimated_sale_price or 0)}} · <b>Total recorded cost:</b> ${{'{:,.0f}'.format(total_cost)}}</p><div class="muted">Provider: {{vehicle.valuation_provider or 'Not calculated'}} · Confidence: {{vehicle.valuation_confidence or 'Not calculated'}}<br>{{vehicle.valuation_source or ''}}</div><div class="actions"><a class="btn" href="{{url_for('vehicle_detail',vehicle_id=vehicle.id)}}">← Back to Vehicle Inventory</a></div></div>
<div class="panel"><h2>BAM Buying Guide & Deal Score</h2><div class="valuation-grid"><div class="valuation-card"><span>Target Buy Price</span><strong>${{'{:,.0f}'.format(target_buy)}}</strong></div><div class="valuation-card"><span>Maximum Recommended Buy</span><strong>${{'{:,.0f}'.format(max_buy)}}</strong></div><div class="valuation-card"><span>Walk-Away Price</span><strong>${{'{:,.0f}'.format(walk_away)}}</strong></div><div class="valuation-card"><span>Expected Profit</span><strong>${{'{:,.0f}'.format(expected_profit)}}</strong></div><div class="valuation-card"><span>Deal Score</span><strong>{{deal_score}}</strong></div></div><div class="notice" style="margin-top:12px">{{deal_message}}</div></div>
{% endblock %}'''
    return render_template_string(template,vehicle=vehicle,google_url=google_url,total_cost=total_cost,target_buy=target_buy,max_buy=max_buy,walk_away=walk_away,expected_profit=expected_profit,deal_score=deal_score,deal_message=deal_message)


@app.route("/vehicles/<int:vehicle_id>/purchase-agreement")
@login_required
def purchase_agreement(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    conn.close()
    if not vehicle:
        return "Vehicle not found", 404
    return render_template("purchase_agreement.html", vehicle=vehicle)


@app.route("/vehicles/<int:vehicle_id>/receipt")
@login_required
def vehicle_receipt(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    sale = conn.execute("SELECT * FROM sales WHERE vehicle_id=?", (vehicle_id,)).fetchone()
    conn.close()
    if not vehicle or not sale:
        return "Receipt is available after a sale is recorded.", 404
    balance_paid = sale["sale_price_inc_gst"] - sale["deposit_amount"] - sale["trade_in_value"]
    return render_template(
        "vehicle_receipt.html",
        vehicle=vehicle,
        sale=sale,
        balance_paid=balance_paid,
    )


@app.route("/vehicles/<int:vehicle_id>/reminders", methods=["POST"])
@login_required
def add_reminder(vehicle_id):
    conn = db()
    conn.execute("""
        INSERT INTO reminders(vehicle_id,reminder_date,reminder_type,title,notes)
        VALUES(?,?,?,?,?)
    """, (
        vehicle_id,
        request.form.get("reminder_date"),
        request.form.get("reminder_type"),
        request.form.get("title"),
        request.form.get("notes"),
    ))
    conn.commit()
    conn.close()
    flash("Reminder added.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.route("/reminders/<int:reminder_id>/complete", methods=["POST"])
@login_required
def complete_reminder(reminder_id):
    conn = db()
    reminder = conn.execute("SELECT * FROM reminders WHERE id=?", (reminder_id,)).fetchone()
    if not reminder:
        conn.close()
        return "Reminder not found", 404
    conn.execute("UPDATE reminders SET completed=1 WHERE id=?", (reminder_id,))
    conn.commit()
    vehicle_id = reminder["vehicle_id"]
    conn.close()
    flash("Reminder completed.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.route("/reminders")
@login_required
def reminders_page():
    conn = db()
    rows = conn.execute("""
        SELECT r.*,v.stock_no,v.make,v.model
        FROM reminders r
        LEFT JOIN vehicles v ON v.id=r.vehicle_id
        ORDER BY r.completed,r.reminder_date,r.id
    """).fetchall()
    conn.close()
    return render_template("reminders.html", rows=rows)


@app.route("/tasks", methods=["GET", "POST"])
@login_required
def tasks_page():
    conn = db()
    if request.method == "POST":
        conn.execute("""
            INSERT INTO tasks(
                vehicle_id,task_date,due_date,title,category,assigned_to,priority,status,notes
            ) VALUES(?,?,?,?,?,?,?,?,?)
        """, (
            int(request.form.get("vehicle_id")) if request.form.get("vehicle_id") else None,
            request.form.get("task_date"),
            request.form.get("due_date"),
            request.form.get("title"),
            request.form.get("category"),
            request.form.get("assigned_to"),
            request.form.get("priority") or "Normal",
            request.form.get("status") or "Open",
            request.form.get("notes"),
        ))
        conn.commit()
        conn.close()
        flash("Task added.", "success")
        return redirect(url_for("tasks_page"))

    rows = conn.execute("""
        SELECT t.*,v.stock_no,v.make,v.model
        FROM tasks t
        LEFT JOIN vehicles v ON v.id=t.vehicle_id
        ORDER BY CASE WHEN t.status='Completed' THEN 1 ELSE 0 END,
                 CASE t.priority WHEN 'Urgent' THEN 1 WHEN 'High' THEN 2 ELSE 3 END,
                 COALESCE(t.due_date,'9999-12-31'),t.id DESC
    """).fetchall()
    vehicles = conn.execute("""
        SELECT id,stock_no,make,model FROM vehicles
        WHERE status NOT IN ('Sold','BER')
        ORDER BY stock_no
    """).fetchall()
    conn.close()
    return render_template("tasks.html", rows=rows, vehicles=vehicles)


@app.route("/tasks/<int:task_id>/complete", methods=["POST"])
@login_required
def complete_task(task_id):
    conn = db()
    conn.execute("UPDATE tasks SET status='Completed' WHERE id=?", (task_id,))
    conn.commit()
    conn.close()
    flash("Task completed.", "success")
    return redirect(request.referrer or url_for("tasks_page"))


@app.route("/vehicles/<int:vehicle_id>/prepare-for-sale", methods=["POST"])
@login_required
def prepare_for_sale(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close()
        return "Vehicle not found", 404

    today_text = date.today().isoformat()
    tasks = [
        ("Workshop", "Complete mechanical inspection", "Barry", "High"),
        ("Compliance", "Confirm PPSR and registration details", "Barry", "High"),
        ("Presentation", "Complete detailing and final clean", "Matt", "Normal"),
        ("Photography", "Upload complete vehicle photo set", "Matt", "Normal"),
        ("Advertising", "Create advertisement and window card", "Barry", "Normal"),
        ("Sales", "Confirm asking price and minimum sale price", "Barry", "High"),
    ]
    for category, title, assigned_to, priority in tasks:
        exists = conn.execute("""
            SELECT id FROM tasks
            WHERE vehicle_id=? AND title=? AND status!='Completed'
        """, (vehicle_id, title)).fetchone()
        if not exists:
            conn.execute("""
                INSERT INTO tasks(vehicle_id,task_date,due_date,title,category,assigned_to,priority,status)
                VALUES(?,?,?,?,?,?,?,'Open')
            """, (
                vehicle_id,
                today_text,
                (date.today() + timedelta(days=7)).isoformat(),
                title,
                category,
                assigned_to,
                priority,
            ))

    conn.execute(
        "UPDATE vehicles SET status='Being Repaired' WHERE id=? AND status='In Stock'",
        (vehicle_id,),
    )
    conn.commit()
    conn.close()
    log_action("Prepare for sale workflow created", "vehicle", vehicle_id, vehicle["stock_no"])
    flash("Prepare-for-sale tasks created.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.route("/finance-calculator", methods=["GET", "POST"])
@login_required
def finance_calculator():
    result = None
    values = {
        "vehicle_price": request.form.get("vehicle_price", ""),
        "deposit": request.form.get("deposit", ""),
        "trade_in": request.form.get("trade_in", ""),
        "annual_rate": request.form.get("annual_rate", ""),
        "term_years": request.form.get("term_years", "5"),
        "balloon": request.form.get("balloon", "0"),
    }
    if request.method == "POST":
        price = float(values["vehicle_price"] or 0)
        deposit = float(values["deposit"] or 0)
        trade_in = float(values["trade_in"] or 0)
        annual_rate = float(values["annual_rate"] or 0) / 100
        years = int(float(values["term_years"] or 0))
        balloon = float(values["balloon"] or 0)
        principal = max(price - deposit - trade_in, 0)
        months = max(years * 12, 1)
        monthly_rate = annual_rate / 12

        financed_before_balloon = max(principal - balloon, 0)
        if monthly_rate > 0:
            monthly_payment = financed_before_balloon * (
                monthly_rate * (1 + monthly_rate) ** months
            ) / ((1 + monthly_rate) ** months - 1)
        else:
            monthly_payment = financed_before_balloon / months

        total_monthly = monthly_payment * months
        total_repaid = total_monthly + balloon
        total_interest = total_repaid - principal
        result = {
            "principal": principal,
            "monthly": monthly_payment,
            "weekly": monthly_payment * 12 / 52,
            "fortnightly": monthly_payment * 12 / 26,
            "total_repaid": total_repaid,
            "total_interest": total_interest,
            "balloon": balloon,
        }

    return render_template("finance_calculator.html", result=result, values=values)


@app.route("/reports/stock-age")
@login_required
def stock_age_report():
    conn = db()
    today_text = date.today().isoformat()
    rows = conn.execute("""
        SELECT v.*,
          CAST(julianday(?) - julianday(COALESCE(v.purchase_date,substr(v.created_at,1,10))) AS INTEGER) AS days_in_stock,
          COALESCE((SELECT SUM(e.cost_inc_gst) FROM expenses e WHERE e.vehicle_id=v.id),0) AS expenses,
          COALESCE((SELECT SUM(CASE WHEN j.actual_cost_inc_gst>0 THEN j.actual_cost_inc_gst ELSE j.estimated_cost END)
                    FROM job_cards j WHERE j.vehicle_id=v.id),0) AS jobs
        FROM vehicles v
        WHERE v.status NOT IN ('Sold','BER')
        ORDER BY days_in_stock DESC
    """, (today_text,)).fetchall()
    conn.close()
    return render_template("stock_age_report.html", rows=rows)


@app.route("/vehicles/<int:vehicle_id>/advertisement-pro")
@login_required
def advertisement_pro(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    photos = conn.execute(
        "SELECT * FROM vehicle_photos WHERE vehicle_id=? ORDER BY id",
        (vehicle_id,),
    ).fetchall()
    conn.close()
    if not vehicle:
        return "Vehicle not found", 404

    title = vehicle["advertisement_title"] or (
        f"{vehicle['year'] or ''} {vehicle['make']} {vehicle['model']} {vehicle['variant'] or ''}".strip()
    )
    description = vehicle["advertisement_description"] or (
        f"{title}\n\n"
        f"• {vehicle['odometer_km'] or 0:,} km\n"
        f"• Registration: {vehicle['registration'] or 'Not listed'}\n"
        f"• Colour: {vehicle['colour'] or 'Not listed'}\n"
        f"• Roadworthy: {vehicle['roadworthy_status'] or 'Not checked'}\n\n"
        f"{vehicle['notes'] or ''}\n\n"
        "Contact BAM Motor Group — Buy • Sell • Trade."
    )
    marketplace = description
    carsales = description + "\n\nStock Number: " + vehicle["stock_no"]
    gumtree = description + "\n\nInspection by appointment."
    return render_template(
        "advertisement_pro.html",
        vehicle=vehicle,
        photos=photos,
        title=title,
        marketplace=marketplace,
        carsales=carsales,
        gumtree=gumtree,
    )

@app.route("/vehicles/<int:vehicle_id>/window-card")
@login_required
def window_card(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    conn.close()
    if not vehicle:
        return "Vehicle not found", 404
    return render_template("window_card.html", vehicle=vehicle)


@app.route("/shipping")
@login_required
def shipping_centre():
    q = request.args.get("q", "").strip()
    status = request.args.get("status", "").strip()
    conn = db()
    sql = """
        SELECT sh.*, p.part_number, p.part_name, p.vehicle_stock_no
        FROM part_shipments sh
        LEFT JOIN parts p ON p.id=sh.part_id
        WHERE 1=1
    """
    params = []
    if q:
        sql += " AND (sh.shipment_number LIKE ? OR sh.customer_name LIKE ? OR sh.tracking_number LIKE ? OR p.part_number LIKE ? OR p.part_name LIKE ?)"
        params.extend([f"%{q}%"] * 5)
    if status:
        sql += " AND sh.status=?"
        params.append(status)
    sql += " ORDER BY sh.id DESC"
    shipments = conn.execute(sql, params).fetchall()
    metrics = conn.execute("""
        SELECT COUNT(*) AS total,
               SUM(CASE WHEN status='Ready to Pack' THEN 1 ELSE 0 END) AS ready,
               SUM(CASE WHEN status='Sent' THEN 1 ELSE 0 END) AS sent,
               SUM(CASE WHEN status='Delivered' THEN 1 ELSE 0 END) AS delivered,
               COALESCE(SUM(freight_charged-shipping_cost),0) AS freight_margin
        FROM part_shipments
    """).fetchone()
    conn.close()
    return render_template("shipping.html", shipments=shipments, metrics=metrics, q=q, status=status)


@app.route("/shipping/new", methods=["GET", "POST"])
@login_required
def shipping_new():
    conn = db()
    parts = conn.execute("""
        SELECT id,part_number,part_name,quantity_on_hand,vehicle_stock_no
        FROM parts WHERE quantity_on_hand>0 ORDER BY part_name
    """).fetchall()
    suggested_number = next_shipment_number(conn)
    selected_part_id = request.args.get("part_id", type=int)

    if request.method == "POST":
        part_id = request.form.get("part_id", type=int)
        quantity = float(request.form.get("quantity") or 1)
        deduct_stock = request.form.get("deduct_stock") == "1"
        if quantity <= 0:
            flash("Shipment quantity must be greater than zero.", "error")
        else:
            part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone() if part_id else None
            if deduct_stock and part and quantity > float(part["quantity_on_hand"] or 0):
                flash("Not enough part stock is available for this shipment.", "error")
            else:
                shipment_number = request.form.get("shipment_number", "").strip() or next_shipment_number(conn)
                cursor = conn.execute("""
                    INSERT INTO part_shipments(
                        shipment_number,part_id,quantity,customer_name,customer_phone,customer_email,
                        address_line,suburb,state,postcode,courier,tracking_number,parcel_weight_kg,
                        parcel_length_cm,parcel_width_cm,parcel_height_cm,shipping_cost,freight_charged,
                        status,date_sent,invoice_included,bubble_wrapped,box_sealed,tracking_sent,
                        stock_adjusted,notes
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    shipment_number, part_id, quantity, request.form.get("customer_name"),
                    request.form.get("customer_phone"), request.form.get("customer_email"),
                    request.form.get("address_line"), request.form.get("suburb"),
                    request.form.get("state"), request.form.get("postcode"),
                    request.form.get("courier"), request.form.get("tracking_number"),
                    float(request.form.get("parcel_weight_kg") or 0),
                    float(request.form.get("parcel_length_cm") or 0),
                    float(request.form.get("parcel_width_cm") or 0),
                    float(request.form.get("parcel_height_cm") or 0),
                    float(request.form.get("shipping_cost") or 0),
                    float(request.form.get("freight_charged") or 0),
                    request.form.get("status") or "Ready to Pack",
                    request.form.get("date_sent") or None,
                    1 if request.form.get("invoice_included") else 0,
                    1 if request.form.get("bubble_wrapped") else 0,
                    1 if request.form.get("box_sealed") else 0,
                    1 if request.form.get("tracking_sent") else 0,
                    1 if deduct_stock and part else 0,
                    request.form.get("notes"),
                ))
                if deduct_stock and part:
                    conn.execute(
                        "UPDATE parts SET quantity_on_hand=quantity_on_hand-? WHERE id=?",
                        (quantity, part_id),
                    )
                conn.commit()
                shipment_id = cursor.lastrowid
                conn.close()
                log_action("Part shipment created", "shipment", shipment_id, shipment_number)
                flash(f"Shipment {shipment_number} created.", "success")
                return redirect(url_for("shipping_detail", shipment_id=shipment_id))

    conn.close()
    return render_template(
        "shipping_form.html", parts=parts, suggested_number=suggested_number,
        selected_part_id=selected_part_id
    )


@app.route("/shipping/<int:shipment_id>", methods=["GET", "POST"])
@login_required
def shipping_detail(shipment_id):
    conn = db()
    if request.method == "POST":
        status = request.form.get("status") or "Ready to Pack"
        date_sent = request.form.get("date_sent") or None
        date_delivered = request.form.get("date_delivered") or None
        conn.execute("""
            UPDATE part_shipments SET
              courier=?,tracking_number=?,status=?,date_sent=?,date_delivered=?,
              invoice_included=?,bubble_wrapped=?,box_sealed=?,tracking_sent=?,notes=?
            WHERE id=?
        """, (
            request.form.get("courier"), request.form.get("tracking_number"), status,
            date_sent, date_delivered,
            1 if request.form.get("invoice_included") else 0,
            1 if request.form.get("bubble_wrapped") else 0,
            1 if request.form.get("box_sealed") else 0,
            1 if request.form.get("tracking_sent") else 0,
            request.form.get("notes"), shipment_id,
        ))
        conn.commit()
        flash("Shipment updated.", "success")
    shipment = conn.execute("""
        SELECT sh.*,p.part_number,p.part_name,p.vehicle_stock_no
        FROM part_shipments sh LEFT JOIN parts p ON p.id=sh.part_id
        WHERE sh.id=?
    """, (shipment_id,)).fetchone()
    conn.close()
    if not shipment:
        return "Shipment not found", 404
    return render_template("shipping_detail.html", shipment=shipment)


@app.route("/shipping/<int:shipment_id>/label")
@login_required
def shipping_label(shipment_id):
    conn = db()
    shipment = conn.execute("""
        SELECT sh.*,p.part_number,p.part_name,p.vehicle_stock_no
        FROM part_shipments sh LEFT JOIN parts p ON p.id=sh.part_id
        WHERE sh.id=?
    """, (shipment_id,)).fetchone()
    conn.close()
    if not shipment:
        return "Shipment not found", 404
    return render_template("shipping_label.html", shipment=shipment)


@app.route("/iphone-access")
@login_required
def iphone_access():
    ip = local_network_ip()
    return render_template(
        "iphone_access.html",
        network_ip=ip,
        mobile_url=url_for("mobile_operations", _external=True),
        local_mobile_url=f"http://{ip}:5000/mobile",
    )



def next_equipment_number(conn=None):
    own_conn = conn is None
    conn = conn or db()
    highest = 0
    for row in conn.execute("SELECT equipment_no FROM equipment WHERE equipment_no LIKE 'EQ-%'").fetchall():
        match = re.search(r"(\d+)$", row["equipment_no"] or "")
        if match:
            highest = max(highest, int(match.group(1)))
    if own_conn:
        conn.close()
    return f"EQ-{highest + 1:05d}"


@app.route("/equipment")
@login_required
def equipment_list():
    q = request.args.get("q", "").strip()
    status = request.args.get("status", "").strip()
    conn = db()
    where = []
    params = []
    if q:
        where.append("(equipment_no LIKE ? OR name LIKE ? OR brand LIKE ? OR model LIKE ? OR serial_number LIKE ? OR category LIKE ? OR location LIKE ? OR assigned_to LIKE ?)")
        params.extend([f"%{q}%"] * 8)
    if status:
        where.append("status=?")
        params.append(status)
    sql = "SELECT * FROM equipment" + ((" WHERE " + " AND ".join(where)) if where else "") + " ORDER BY name, equipment_no"
    rows = conn.execute(sql, params).fetchall()
    metrics = conn.execute("""
        SELECT COUNT(*) total,
               COALESCE(SUM(current_value),0) total_value,
               SUM(CASE WHEN status='Checked Out' THEN 1 ELSE 0 END) checked_out,
               SUM(CASE WHEN next_service_date IS NOT NULL AND next_service_date!='' AND next_service_date<=date('now','+30 day') THEN 1 ELSE 0 END) due_service
        FROM equipment
    """).fetchone()
    conn.close()
    return render_template("equipment.html", equipment=rows, metrics=metrics, q=q, selected_status=status)


@app.route("/equipment/new", methods=["GET", "POST"])
@login_required
def equipment_new():
    if request.method == "POST":
        try:
            photo = save_upload(request.files.get("photo"))
            receipt = save_upload(request.files.get("receipt"))
            conn = db()
            equipment_no = request.form.get("equipment_no", "").strip() or next_equipment_number(conn)
            cursor = conn.execute("""
                INSERT INTO equipment(
                    equipment_no,name,category,brand,model,serial_number,purchase_date,purchase_price,current_value,
                    supplier,warranty_expiry,location,assigned_to,condition,status,next_service_date,calibration_due,
                    test_tag_due,photo_filename,receipt_filename,notes
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                equipment_no, request.form.get("name", "").strip(), request.form.get("category"),
                request.form.get("brand"), request.form.get("model"), request.form.get("serial_number"),
                request.form.get("purchase_date"), float(request.form.get("purchase_price") or 0),
                float(request.form.get("current_value") or 0), request.form.get("supplier"),
                request.form.get("warranty_expiry"), request.form.get("location"), request.form.get("assigned_to"),
                request.form.get("condition") or "Good", request.form.get("status") or "Available",
                request.form.get("next_service_date"), request.form.get("calibration_due"), request.form.get("test_tag_due"),
                photo, receipt, request.form.get("notes")
            ))
            conn.execute("INSERT INTO equipment_history(equipment_id,action,assigned_to,location,condition,notes,user_name) VALUES(?,?,?,?,?,?,?)",
                         (cursor.lastrowid,"Equipment added",request.form.get("assigned_to"),request.form.get("location"),request.form.get("condition"),request.form.get("notes"),session.get("display_name")))
            conn.commit(); conn.close()
            flash("Equipment added.", "success")
            return redirect(url_for("equipment_detail", equipment_id=cursor.lastrowid))
        except (ValueError, sqlite3.IntegrityError) as exc:
            flash(str(exc), "error")
    conn=db(); suggested=next_equipment_number(conn); conn.close()
    return render_template("equipment_form.html", item=None, suggested_no=suggested)


@app.route("/equipment/<int:equipment_id>")
@login_required
def equipment_detail(equipment_id):
    conn=db()
    item=conn.execute("SELECT * FROM equipment WHERE id=?",(equipment_id,)).fetchone()
    history=conn.execute("SELECT * FROM equipment_history WHERE equipment_id=? ORDER BY action_date DESC,id DESC",(equipment_id,)).fetchall()
    conn.close()
    if not item: return "Equipment not found",404
    return render_template("equipment_detail.html", item=item, history=history)


@app.route("/equipment/<int:equipment_id>/edit", methods=["GET","POST"])
@login_required
def equipment_edit(equipment_id):
    conn=db(); item=conn.execute("SELECT * FROM equipment WHERE id=?",(equipment_id,)).fetchone()
    if not item:
        conn.close(); return "Equipment not found",404
    if request.method == "POST":
        try:
            photo=save_upload(request.files.get("photo")) or item["photo_filename"]
            receipt=save_upload(request.files.get("receipt")) or item["receipt_filename"]
            conn.execute("""UPDATE equipment SET equipment_no=?,name=?,category=?,brand=?,model=?,serial_number=?,purchase_date=?,purchase_price=?,current_value=?,supplier=?,warranty_expiry=?,location=?,assigned_to=?,condition=?,status=?,next_service_date=?,calibration_due=?,test_tag_due=?,photo_filename=?,receipt_filename=?,notes=? WHERE id=?""",(
                request.form.get("equipment_no"),request.form.get("name"),request.form.get("category"),request.form.get("brand"),request.form.get("model"),request.form.get("serial_number"),request.form.get("purchase_date"),float(request.form.get("purchase_price") or 0),float(request.form.get("current_value") or 0),request.form.get("supplier"),request.form.get("warranty_expiry"),request.form.get("location"),request.form.get("assigned_to"),request.form.get("condition"),request.form.get("status"),request.form.get("next_service_date"),request.form.get("calibration_due"),request.form.get("test_tag_due"),photo,receipt,request.form.get("notes"),equipment_id))
            conn.execute("INSERT INTO equipment_history(equipment_id,action,assigned_to,location,condition,notes,user_name) VALUES(?,?,?,?,?,?,?)",(equipment_id,"Equipment updated",request.form.get("assigned_to"),request.form.get("location"),request.form.get("condition"),"Details updated",session.get("display_name")))
            conn.commit(); conn.close(); flash("Equipment updated.","success")
            return redirect(url_for("equipment_detail",equipment_id=equipment_id))
        except (ValueError,sqlite3.IntegrityError) as exc:
            flash(str(exc),"error")
    conn.close()
    return render_template("equipment_form.html",item=item,suggested_no=item["equipment_no"])


@app.route("/equipment/<int:equipment_id>/movement", methods=["POST"])
@login_required
def equipment_movement(equipment_id):
    action=request.form.get("action") or "Updated"
    status="Checked Out" if action=="Checked Out" else ("Available" if action=="Returned" else request.form.get("status") or "Available")
    conn=db()
    conn.execute("UPDATE equipment SET assigned_to=?,location=?,condition=?,status=? WHERE id=?",(request.form.get("assigned_to"),request.form.get("location"),request.form.get("condition") or "Good",status,equipment_id))
    conn.execute("INSERT INTO equipment_history(equipment_id,action,assigned_to,location,condition,notes,user_name) VALUES(?,?,?,?,?,?,?)",(equipment_id,action,request.form.get("assigned_to"),request.form.get("location"),request.form.get("condition"),request.form.get("notes"),session.get("display_name")))
    conn.commit(); conn.close(); flash(f"Equipment {action.lower()}.","success")
    return redirect(url_for("equipment_detail",equipment_id=equipment_id))


@app.route("/users")
@owner_required
def users_page():
    conn=db(); users=conn.execute("SELECT * FROM users ORDER BY display_name,username").fetchall(); conn.close()
    return render_template("users.html", users=users)


@app.route("/users/new", methods=["GET","POST"])
@owner_required
def user_new():
    if request.method == "POST":
        username=(request.form.get("username") or "").strip().lower()
        password=request.form.get("password") or ""
        display_name=(request.form.get("display_name") or "").strip()
        role=request.form.get("role") or "staff"
        if not username or not display_name or len(password)<8:
            flash("Username, display name and a password of at least 8 characters are required.","error")
        else:
            try:
                conn=db(); conn.execute("INSERT INTO users(username,password_hash,display_name,role,is_active) VALUES(?,?,?,?,1)",(username,generate_password_hash(password),display_name,role)); conn.commit(); conn.close()
                flash("User created.","success"); return redirect(url_for("users_page"))
            except sqlite3.IntegrityError:
                flash("That username is already in use.","error")
    return render_template("user_form.html", user=None)


@app.route("/users/<int:user_id>/edit", methods=["GET","POST"])
@owner_required
def user_edit(user_id):
    conn=db(); user=conn.execute("SELECT * FROM users WHERE id=?",(user_id,)).fetchone()
    if not user: conn.close(); return "User not found",404
    if request.method == "POST":
        username=(request.form.get("username") or "").strip().lower(); display_name=(request.form.get("display_name") or "").strip(); role=request.form.get("role") or "staff"; active=1 if request.form.get("is_active") else 0; password=request.form.get("password") or ""
        if user_id==session.get("user_id") and not active:
            conn.close(); flash("You cannot disable your own account.","error"); return redirect(url_for("user_edit",user_id=user_id))
        try:
            if password:
                if len(password)<8: raise ValueError("Password must be at least 8 characters.")
                conn.execute("UPDATE users SET username=?,display_name=?,role=?,is_active=?,password_hash=? WHERE id=?",(username,display_name,role,active,generate_password_hash(password),user_id))
            else:
                conn.execute("UPDATE users SET username=?,display_name=?,role=?,is_active=? WHERE id=?",(username,display_name,role,active,user_id))
            conn.commit(); conn.close(); flash("User updated.","success"); return redirect(url_for("users_page"))
        except (sqlite3.IntegrityError,ValueError) as exc:
            conn.close(); flash(str(exc),"error")
    else: conn.close()
    return render_template("user_form.html", user=user)


def _workshop_provider_vehicle_url(vehicle):
    if not WORKSHOP_PORTAL_URL:
        return ""
    replacements = {
        "vin": urllib.parse.quote(str(vehicle["vin"] or "")),
        "make": urllib.parse.quote(str(vehicle["make"] or "")),
        "model": urllib.parse.quote(str(vehicle["model"] or "")),
        "year": urllib.parse.quote(str(vehicle["year"] or "")),
        "stock_no": urllib.parse.quote(str(vehicle["stock_no"] or "")),
    }
    url = WORKSHOP_PORTAL_URL
    for key, value in replacements.items():
        url = url.replace("{" + key + "}", value)
    return url


@app.route("/workshop-intelligence")
@login_required
def workshop_centre():
    q = (request.args.get("q") or "").strip()
    conn = db()
    params = []
    where = ""
    if q:
        like = f"%{q}%"
        where = "WHERE stock_no LIKE ? OR vin LIKE ? OR registration LIKE ? OR make LIKE ? OR model LIKE ?"
        params = [like] * 5
    vehicles = conn.execute(f"""
        SELECT id,stock_no,year,make,model,variant,vin,registration,status
        FROM vehicles {where}
        ORDER BY id DESC LIMIT 50
    """, params).fetchall()
    recent_jobs = conn.execute("""
        SELECT j.*,v.stock_no,v.make,v.model
        FROM job_cards j JOIN vehicles v ON v.id=j.vehicle_id
        ORDER BY j.id DESC LIMIT 12
    """).fetchall()
    conn.close()
    return render_template("workshop_centre.html", vehicles=vehicles, q=q, recent_jobs=recent_jobs,
                           default_labour_rate=DEFAULT_LABOUR_RATE)


@app.route("/vehicles/<int:vehicle_id>/workshop-intelligence")
@login_required
def workshop_intelligence(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close(); return "Vehicle not found", 404
    operations = conn.execute("""
        SELECT o.*,p.part_number,p.part_name
        FROM workshop_operations o
        LEFT JOIN parts p ON p.id=o.part_id
        WHERE o.vehicle_id=? OR (
            o.vehicle_id IS NULL AND LOWER(COALESCE(o.make,''))=LOWER(COALESCE(?,''))
            AND LOWER(COALESCE(o.model,''))=LOWER(COALESCE(?,''))
            AND (o.year IS NULL OR o.year=0 OR o.year=?)
        )
        ORDER BY o.system_name,o.operation_name,o.id DESC
    """, (vehicle_id, vehicle["make"], vehicle["model"], vehicle["year"] or 0)).fetchall()
    refs = conn.execute("SELECT * FROM workshop_references WHERE vehicle_id=? ORDER BY reference_type,title", (vehicle_id,)).fetchall()
    parts = conn.execute("SELECT id,part_number,part_name,manufacturer_part_no,quantity_on_hand FROM parts WHERE quantity_on_hand>0 ORDER BY part_name").fetchall()
    jobs = conn.execute("SELECT * FROM job_cards WHERE vehicle_id=? ORDER BY id DESC LIMIT 20", (vehicle_id,)).fetchall()
    conn.close()
    specs = {}
    try:
        specs = json.loads(vehicle["vin_decode_json"] or "{}") if vehicle["vin_decode_json"] else {}
    except Exception:
        specs = {}
    manual_search_phrase = " ".join(str(x).strip() for x in [vehicle["year"], vehicle["make"], vehicle["model"], vehicle["variant"], "workshop service repair manual"] if x)
    return render_template("workshop_intelligence.html", vehicle=vehicle, operations=operations, references=refs,
                           parts=parts, jobs=jobs, provider_name=WORKSHOP_PROVIDER_NAME,
                           provider_url=_workshop_provider_vehicle_url(vehicle), default_labour_rate=DEFAULT_LABOUR_RATE,
                           decoded_specs=specs, manual_search_phrase=manual_search_phrase)


@app.route("/vehicles/<int:vehicle_id>/workshop-operation", methods=["POST"])
@login_required
def workshop_operation_add(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close(); return "Vehicle not found", 404
    try:
        hours = float(request.form.get("labour_hours") or 0)
        rate = float(request.form.get("labour_rate") or DEFAULT_LABOUR_RATE)
        if hours < 0 or rate < 0:
            raise ValueError("Labour hours and labour rate cannot be negative.")
        part_id = int(request.form.get("part_id")) if request.form.get("part_id") else None
        operation_name = (request.form.get("operation_name") or "").strip()
        if not operation_name and part_id:
            part = conn.execute("SELECT part_name FROM parts WHERE id=?", (part_id,)).fetchone()
            operation_name = f"Fit {part['part_name']}" if part else "Fit part"
        if not operation_name:
            raise ValueError("Enter an operation name.")
        reusable = request.form.get("scope") == "library"
        conn.execute("""
            INSERT INTO workshop_operations(vehicle_id,part_id,operation_code,system_name,operation_name,labour_hours,labour_rate,
                source_name,source_reference,procedure_url,notes,make,model,year,engine_hint,created_by)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (None if reusable else vehicle_id, part_id, request.form.get("operation_code"), request.form.get("system_name"), operation_name,
              hours, rate, request.form.get("source_name"), request.form.get("source_reference"), request.form.get("procedure_url"),
              request.form.get("notes"), vehicle["make"], vehicle["model"], vehicle["year"],
              vehicle["engine_model"] or vehicle["engine_make"] or vehicle["engine_displacement_l"], session.get("display_name")))
        conn.commit()
        log_action("Workshop labour operation added", "vehicle", vehicle_id, f"{operation_name} — {hours:.2f} h")
        flash("Workshop labour operation saved.", "success")
    except (ValueError, sqlite3.Error) as exc:
        conn.rollback(); flash(str(exc), "error")
    finally:
        conn.close()
    return redirect(url_for("workshop_intelligence", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/workshop-operation/<int:operation_id>/job", methods=["POST"])
@login_required
def workshop_operation_to_job(vehicle_id, operation_id):
    conn = db()
    op = conn.execute("SELECT * FROM workshop_operations WHERE id=?", (operation_id,)).fetchone()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not op or not vehicle:
        conn.close(); return "Workshop operation or vehicle not found", 404
    hours = float(op["labour_hours"] or 0)
    rate = float(op["labour_rate"] or DEFAULT_LABOUR_RATE)
    estimate = round(hours * rate, 2)
    notes = "\n".join(x for x in [
        f"Labour source: {op['source_name']}" if op["source_name"] else "",
        f"Reference: {op['source_reference']}" if op["source_reference"] else "",
        op["notes"] or "",
    ] if x)
    cur = conn.execute("""
        INSERT INTO job_cards(vehicle_id,job_date,category,description,supplier,paid_by,estimated_cost,actual_cost_inc_gst,
            gst_amount,status,notes,labour_operation_id,labour_hours,labour_rate,labour_source,labour_code,procedure_url)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (vehicle_id, date.today().isoformat(), op["system_name"] or "Mechanical", op["operation_name"], "", "Business", estimate, 0, 0,
          "Open", notes, op["id"], hours, rate, op["source_name"], op["operation_code"], op["procedure_url"]))
    job_id = cur.lastrowid
    conn.execute("INSERT INTO job_card_history(job_card_id,old_status,new_status,changed_by,note) VALUES(?,?,?,?,?)",
                 (job_id, None, "Open", session.get("display_name"), "Created from Workshop Intelligence labour operation"))
    conn.commit(); conn.close()
    log_action("Workshop operation added to job card", "vehicle", vehicle_id, f"Job #{job_id}; {hours:.2f} h @ ${rate:.2f}")
    flash(f"Job card #{job_id} created with {hours:.2f} labour hours.", "success")
    return redirect(url_for("vehicle_detail", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/workshop-operation/<int:operation_id>/delete", methods=["POST"])
@login_required
def workshop_operation_delete(vehicle_id, operation_id):
    conn = db()
    conn.execute("DELETE FROM workshop_operations WHERE id=? AND (vehicle_id=? OR vehicle_id IS NULL)", (operation_id, vehicle_id))
    conn.commit(); conn.close()
    flash("Workshop labour operation removed.", "success")
    return redirect(url_for("workshop_intelligence", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/workshop-reference", methods=["POST"])
@login_required
def workshop_reference_add(vehicle_id):
    title = (request.form.get("title") or "").strip()
    if not title:
        flash("Reference title is required.", "error")
        return redirect(url_for("workshop_intelligence", vehicle_id=vehicle_id))
    conn = db()
    conn.execute("""
        INSERT INTO workshop_references(vehicle_id,reference_type,title,provider_name,reference_url,reference_code,notes,created_by)
        VALUES(?,?,?,?,?,?,?,?)
    """, (vehicle_id, request.form.get("reference_type") or "Workshop Manual", title, request.form.get("provider_name"),
          request.form.get("reference_url"), request.form.get("reference_code"), request.form.get("notes"), session.get("display_name")))
    conn.commit(); conn.close()
    flash("Workshop technical reference saved.", "success")
    return redirect(url_for("workshop_intelligence", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/workshop-reference/<int:reference_id>/delete", methods=["POST"])
@login_required
def workshop_reference_delete(vehicle_id, reference_id):
    conn = db(); conn.execute("DELETE FROM workshop_references WHERE id=? AND vehicle_id=?", (reference_id, vehicle_id)); conn.commit(); conn.close()
    flash("Workshop reference removed.", "success")
    return redirect(url_for("workshop_intelligence", vehicle_id=vehicle_id))


@app.route("/vehicles/<int:vehicle_id>/workshop-labour-import", methods=["POST"])
@login_required
def workshop_labour_import(vehicle_id):
    upload = request.files.get("labour_csv")
    if not upload or not upload.filename:
        flash("Choose a CSV file first.", "error")
        return redirect(url_for("workshop_intelligence", vehicle_id=vehicle_id))
    conn = db(); vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close(); return "Vehicle not found", 404
    try:
        raw = upload.read().decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(raw))
        required = {"operation_name", "labour_hours"}
        if not reader.fieldnames or not required.issubset({x.strip() for x in reader.fieldnames}):
            raise ValueError("CSV must include operation_name and labour_hours columns.")
        imported = 0
        for row in reader:
            name = (row.get("operation_name") or "").strip()
            if not name: continue
            hours = float(row.get("labour_hours") or 0)
            rate = float(row.get("labour_rate") or DEFAULT_LABOUR_RATE)
            conn.execute("""
                INSERT INTO workshop_operations(vehicle_id,operation_code,system_name,operation_name,labour_hours,labour_rate,
                    source_name,source_reference,procedure_url,notes,make,model,year,engine_hint,created_by)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (vehicle_id, row.get("operation_code"), row.get("system_name"), name, hours, rate, row.get("source_name"),
                  row.get("source_reference"), row.get("procedure_url"), row.get("notes"), vehicle["make"], vehicle["model"],
                  vehicle["year"], vehicle["engine_model"] or vehicle["engine_make"] or vehicle["engine_displacement_l"], session.get("display_name")))
            imported += 1
        conn.commit(); flash(f"Imported {imported} workshop labour operation(s).", "success")
    except (UnicodeDecodeError, ValueError, sqlite3.Error) as exc:
        conn.rollback(); flash(f"Labour import failed: {exc}", "error")
    finally:
        conn.close()
    return redirect(url_for("workshop_intelligence", vehicle_id=vehicle_id))


@app.route("/email-centre", methods=["GET", "POST"])
@login_required
def email_centre():
    conn = db()
    if request.method == "POST":
        subject = (request.form.get("subject") or "").strip()
        if not subject:
            flash("Email subject is required.", "error")
        else:
            try:
                conn.execute("""
                    INSERT INTO email_messages(direction,message_date,sender,recipient,subject,body_excerpt,status,
                        follow_up_date,vehicle_id,contact_id,external_reference,created_by)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    request.form.get("direction") or "Incoming",
                    request.form.get("message_date") or date.today().isoformat(),
                    request.form.get("sender"), request.form.get("recipient"), subject,
                    request.form.get("body_excerpt"), request.form.get("status") or "Unread",
                    request.form.get("follow_up_date"),
                    int(request.form.get("vehicle_id")) if request.form.get("vehicle_id") else None,
                    int(request.form.get("contact_id")) if request.form.get("contact_id") else None,
                    request.form.get("external_reference"), session.get("display_name")
                ))
                conn.commit()
                flash("Email correspondence saved.", "success")
                return redirect(url_for("email_centre"))
            except (ValueError, sqlite3.Error) as exc:
                conn.rollback(); flash(f"Could not save email: {exc}", "error")
    q = (request.args.get("q") or "").strip()
    status = (request.args.get("status") or "").strip()
    params=[]; where=[]
    if q:
        like=f"%{q}%"; where.append("(em.subject LIKE ? OR em.sender LIKE ? OR em.recipient LIKE ? OR em.body_excerpt LIKE ? OR v.stock_no LIKE ? OR c.name LIKE ?)"); params += [like]*6
    if status:
        where.append("em.status=?"); params.append(status)
    sql="""
        SELECT em.*, v.stock_no, v.make, v.model, c.name AS contact_name
        FROM email_messages em
        LEFT JOIN vehicles v ON v.id=em.vehicle_id
        LEFT JOIN contacts c ON c.id=em.contact_id
    """
    if where: sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY COALESCE(em.message_date, substr(em.created_at,1,10)) DESC, em.id DESC LIMIT 250"
    messages=conn.execute(sql, params).fetchall()
    vehicles=conn.execute("SELECT id,stock_no,year,make,model FROM vehicles ORDER BY stock_no DESC").fetchall()
    contacts=conn.execute("SELECT id,name,email FROM contacts ORDER BY name").fetchall()
    metrics=conn.execute("SELECT COUNT(*) total, SUM(CASE WHEN status='Unread' THEN 1 ELSE 0 END) unread, SUM(CASE WHEN status='Follow Up' THEN 1 ELSE 0 END) follow_up FROM email_messages").fetchone()
    conn.close()
    return render_template("email_centre.html", messages=messages, vehicles=vehicles, contacts=contacts, metrics=metrics,
                           email_webmail_url=EMAIL_WEBMAIL_URL, email_address=EMAIL_ADDRESS, q=q, status_filter=status,
                           today=date.today().isoformat())

@app.route("/email-centre/compose", methods=["GET", "POST"])
@login_required
def compose_email():
    if request.method == "POST":
        recipient = (request.form.get("recipient") or "").strip()
        subject = (request.form.get("subject") or "").strip()
        body = (request.form.get("body") or "").strip()

        if not recipient or not subject:
            flash("Recipient and subject are required.", "error")
            return render_template(
                "compose_email.html",
                email_address=EMAIL_ADDRESS,
                recipient=recipient,
                subject=subject,
                body=body,
            )

        if not all([
            EMAIL_ADDRESS,
            EMAIL_PASSWORD,
            EMAIL_SMTP_SERVER,
            EMAIL_SMTP_PORT,
        ]):
            flash("Email sending is not fully configured.", "error")
            return redirect(url_for("email_centre"))

        try:
            msg = EmailMessage()
            msg["From"] = EMAIL_ADDRESS
            msg["To"] = recipient
            msg["Subject"] = subject
            msg.set_content(body)

            with smtplib.SMTP_SSL(
                EMAIL_SMTP_SERVER,
                EMAIL_SMTP_PORT,
                timeout=30,
            ) as smtp:
                smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                smtp.send_message(msg)

            flash("Email sent successfully.", "success")
            return redirect(url_for("email_centre"))

        except Exception as exc:
            flash(f"Could not send email: {exc}", "error")

    return render_template(
        "compose_email.html",
        email_address=EMAIL_ADDRESS,
    )
@app.route("/email-centre/<int:message_id>/status", methods=["POST"])
@login_required
def email_message_status(message_id):
    new_status = request.form.get("status") or "Read"
    if new_status not in {"Unread","Read","Follow Up","Completed"}: new_status="Read"
    conn=db(); conn.execute("UPDATE email_messages SET status=? WHERE id=?", (new_status,message_id)); conn.commit(); conn.close()
    flash("Email status updated.", "success")
    return redirect(request.referrer or url_for("centre"))


@app.route("/email-centre/<int:message_id>/delete", methods=["POST"])
@login_required
def email_message_delete(message_id):
    conn=db(); conn.execute("DELETE FROM email_messages WHERE id=?", (message_id,)); conn.commit(); conn.close()
    flash("Email record removed.", "success")
    return redirect(url_for("email_centre"))


@app.route("/vehicles/<int:vehicle_id>/emanual-link", methods=["POST"])
@login_required
def emanual_link_add(vehicle_id):
    link=(request.form.get("reference_url") or "").strip()
    title=(request.form.get("title") or "eManualOnline Workshop Manual").strip()
    if not link:
        flash("Paste the purchased eManualOnline manual link first.", "error")
        return redirect(url_for("workshop_intelligence", vehicle_id=vehicle_id))
    conn=db()
    vehicle=conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close(); return "Vehicle not found",404
    conn.execute("""
        INSERT INTO workshop_references(vehicle_id,reference_type,title,provider_name,reference_url,reference_code,notes,created_by)
        VALUES(?,?,?,?,?,?,?,?)
    """, (vehicle_id,"Workshop Manual",title,"eManualOnline",link,vehicle["vin"],"Saved eManualOnline vehicle manual",session.get("display_name")))
    conn.commit(); conn.close()
    flash("eManualOnline manual linked to this vehicle.", "success")
    return redirect(url_for("workshop_intelligence", vehicle_id=vehicle_id))



# ---------------------------------------------------------------------------
# Version 24 - Intelligent Dealer Platform
# Safe local Dealer Copilot + Executive Dashboard.
# The Copilot only runs predefined, read-only queries against BAM's own data.
# ---------------------------------------------------------------------------
def _copilot_money(value):
    try:
        return f"${float(value or 0):,.2f}"
    except Exception:
        return "$0.00"


def dealer_copilot_answer(conn, question):
    q = (question or "").strip()
    low = q.lower()
    result = {"question": q, "summary": "", "items": [], "kind": "general", "tips": []}
    if not q:
        result["summary"] = "Ask about vehicles, parts, workshop jobs, customers, invoices, reminders or business performance."
        return result

    # Exact stock / VIN / registration lookup first.
    vehicle = conn.execute("""
        SELECT id,stock_no,year,make,model,status,vin,registration,asking_price,purchase_price_inc_gst
        FROM vehicles
        WHERE UPPER(COALESCE(stock_no,''))=UPPER(?) OR UPPER(COALESCE(vin,''))=UPPER(?) OR UPPER(COALESCE(registration,''))=UPPER(?)
        LIMIT 1
    """, (q,q,q)).fetchone()
    if vehicle:
        result["kind"] = "vehicle"
        result["summary"] = f"Found {vehicle['stock_no']} — {vehicle['year'] or ''} {vehicle['make']} {vehicle['model']}."
        result["items"] = [{"title": f"{vehicle['stock_no']} — {vehicle['year'] or ''} {vehicle['make']} {vehicle['model']}", "subtitle": f"{vehicle['status']} · VIN {vehicle['vin'] or 'not recorded'} · Rego {vehicle['registration'] or 'not recorded'}", "amount": _copilot_money(vehicle['asking_price']), "url": url_for('vehicle_detail', vehicle_id=vehicle['id'])}]
        return result

    if ("profit" in low or "profitable" in low) and ("vehicle" in low or "car" in low or "stock" in low):
        rows = conn.execute("""
            SELECT v.id,v.stock_no,v.year,v.make,v.model,
              COALESCE((SELECT s.sale_price_inc_gst FROM sales s WHERE s.vehicle_id=v.id),0)
              + COALESCE((SELECT SUM(ps.sale_price) FROM part_sales ps JOIN parts p ON p.id=ps.part_id WHERE p.vehicle_id=v.id),0)
              + COALESCE(v.shell_sale_price,0) AS revenue,
              COALESCE(v.purchase_price_inc_gst,0)
              + COALESCE((SELECT SUM(e.cost_inc_gst) FROM expenses e WHERE e.vehicle_id=v.id),0)
              + COALESCE((SELECT SUM(CASE WHEN j.actual_cost_inc_gst>0 THEN j.actual_cost_inc_gst ELSE j.estimated_cost END) FROM job_cards j WHERE j.vehicle_id=v.id),0)
              + COALESCE((SELECT SUM(se.cost_inc_gst) FROM service_entries se WHERE se.vehicle_id=v.id),0) AS invested
            FROM vehicles v
            ORDER BY (revenue-invested) DESC
            LIMIT 10
        """).fetchall()
        result["kind"] = "profit"
        result["summary"] = "Most profitable vehicles based on recorded vehicle sales, donor-part sales and shell revenue less recorded investment and costs."
        for r in rows:
            profit=float(r['revenue'] or 0)-float(r['invested'] or 0)
            result["items"].append({"title": f"{r['stock_no']} — {r['year'] or ''} {r['make']} {r['model']}", "subtitle": f"Revenue {_copilot_money(r['revenue'])} · Invested {_copilot_money(r['invested'])}", "amount": _copilot_money(profit), "url": url_for('vehicle_detail', vehicle_id=r['id'])})
        return result

    days_match = re.search(r"(\d+)\s*day", low)
    days = int(days_match.group(1)) if days_match else 90
    if ("sitting" in low or "stock age" in low or "days in stock" in low or "old stock" in low):
        today_text=date.today().isoformat()
        rows=conn.execute("""
            SELECT id,stock_no,year,make,model,status,purchase_date,
                   CAST(julianday(?) - julianday(COALESCE(purchase_date,substr(created_at,1,10))) AS INTEGER) AS days_in_stock
            FROM vehicles
            WHERE status NOT IN ('Sold','BER')
              AND (julianday(?) - julianday(COALESCE(purchase_date,substr(created_at,1,10)))) >= ?
            ORDER BY days_in_stock DESC
            LIMIT 30
        """,(today_text,today_text,days)).fetchall()
        result["kind"]="stock_age"
        result["summary"]=f"{len(rows)} active vehicle(s) have been in stock for at least {days} days."
        result["items"]=[{"title":f"{r['stock_no']} — {r['year'] or ''} {r['make']} {r['model']}","subtitle":f"{r['days_in_stock']} days · {r['status']}","amount":"","url":url_for('vehicle_detail',vehicle_id=r['id'])} for r in rows]
        return result

    if "photo" in low and ("missing" in low or "need" in low or "without" in low):
        rows=conn.execute("""
            SELECT id,stock_no,year,make,model,status FROM vehicles
            WHERE status NOT IN ('Sold','BER') AND COALESCE(photo_filename,'')=''
              AND NOT EXISTS (SELECT 1 FROM vehicle_photos p WHERE p.vehicle_id=vehicles.id)
            ORDER BY id DESC LIMIT 30
        """).fetchall()
        result["kind"]="photos"; result["summary"]=f"{len(rows)} active vehicle(s) are missing photos."
        result["items"]=[{"title":f"{r['stock_no']} — {r['year'] or ''} {r['make']} {r['model']}","subtitle":r['status'],"amount":"","url":url_for('vehicle_detail',vehicle_id=r['id'])} for r in rows]
        return result

    if "workshop" in low or "job" in low:
        overdue = "overdue" in low
        sql="""SELECT j.id,j.vehicle_id,j.description,j.status,j.job_date,v.stock_no,v.make,v.model
               FROM job_cards j LEFT JOIN vehicles v ON v.id=j.vehicle_id
               WHERE COALESCE(j.status,'Open')!='Completed'"""
        params=[]
        if overdue:
            sql += " AND COALESCE(j.job_date,'9999-12-31') < ?"; params.append(date.today().isoformat())
        sql += " ORDER BY j.job_date,j.id DESC LIMIT 30"
        rows=conn.execute(sql,params).fetchall()
        result["kind"]="workshop"; result["summary"]=("Overdue" if overdue else "Active")+f" workshop jobs: {len(rows)}."
        result["items"]=[{"title":r['description'] or 'Workshop job',"subtitle":f"{r['stock_no'] or 'No vehicle'} — {r['make'] or ''} {r['model'] or ''} · {r['status']}","amount":"","url":url_for('vehicle_detail',vehicle_id=r['vehicle_id']) if r['vehicle_id'] else url_for('workshop_centre')} for r in rows]
        return result

    if "invoice" in low and ("draft" in low or "outstanding" in low or "unpaid" in low):
        rows=conn.execute("""
            SELECT s.vehicle_id,s.invoice_number,s.invoice_status,s.buyer_name,s.sale_price_inc_gst,v.stock_no,v.make,v.model
            FROM sales s JOIN vehicles v ON v.id=s.vehicle_id
            WHERE COALESCE(s.invoice_status,'Draft')!='Paid'
            ORDER BY s.sale_date DESC LIMIT 30
        """).fetchall()
        result["kind"]="invoice"; result["summary"]=f"{len(rows)} invoice(s) are not marked Paid."
        result["items"]=[{"title":r['invoice_number'] or 'Draft invoice',"subtitle":f"{r['buyer_name'] or 'Buyer'} · {r['stock_no']} {r['make']} {r['model']} · {r['invoice_status'] or 'Draft'}","amount":_copilot_money(r['sale_price_inc_gst']),"url":url_for('sale_invoice',vehicle_id=r['vehicle_id'])} for r in rows]
        return result

    if "reminder" in low or "task" in low or "attention" in low:
        rows=conn.execute("""
            SELECT r.id,r.vehicle_id,r.reminder_date,r.title,r.notes,v.stock_no,v.make,v.model
            FROM reminders r LEFT JOIN vehicles v ON v.id=r.vehicle_id
            WHERE r.completed=0
            ORDER BY COALESCE(r.reminder_date,'9999-12-31'),r.id LIMIT 30
        """).fetchall()
        result["kind"]="reminders"; result["summary"]=f"{len(rows)} open reminder(s) need attention."
        result["items"]=[{"title":r['title'],"subtitle":f"{r['stock_no'] or ''} {r['make'] or ''} {r['model'] or ''} · Due {r['reminder_date'] or 'not set'}","amount":"","url":url_for('vehicle_detail',vehicle_id=r['vehicle_id']) if r['vehicle_id'] else url_for('reminders_page')} for r in rows]
        return result

    if "part" in low or "engine" in low or "gearbox" in low or "transmission" in low:
        stop={"show","me","all","find","parts","part","in","stock","the","a","an","for","from","do","we","have"}
        terms=[t for t in re.findall(r"[a-z0-9-]+",low) if t not in stop and len(t)>1]
        search="%"+"%".join(terms[:4])+"%" if terms else "%"
        rows=conn.execute("""
            SELECT id,part_number,part_name,category,quantity_on_hand,selling_price,storage_location,make,model,vehicle_stock_no,status
            FROM parts
            WHERE (LOWER(COALESCE(part_name,'')) LIKE ? OR LOWER(COALESCE(category,'')) LIKE ? OR LOWER(COALESCE(make,'')) LIKE ? OR LOWER(COALESCE(model,'')) LIKE ? OR LOWER(COALESCE(fitment,'')) LIKE ? OR LOWER(COALESCE(manufacturer_part_no,'')) LIKE ?)
            ORDER BY quantity_on_hand DESC,id DESC LIMIT 40
        """,(search,search,search,search,search,search)).fetchall()
        result["kind"]="parts"; result["summary"]=f"Found {len(rows)} matching part(s)."
        result["items"]=[{"title":f"{r['part_number'] or ''} — {r['part_name']}","subtitle":f"{r['make'] or ''} {r['model'] or ''} · Qty {r['quantity_on_hand'] or 0} · {r['storage_location'] or 'No location'} · {r['status'] or ''}","amount":_copilot_money(r['selling_price']),"url":url_for('part_detail',part_id=r['id'])} for r in rows]
        return result

    if any(k in low for k in ("sales this month","revenue this month","monthly sales","month profit","monthly profit")):
        month=date.today().strftime('%Y-%m')
        vehicle_rev=float(conn.execute("SELECT COALESCE(SUM(sale_price_inc_gst),0) v FROM sales WHERE substr(sale_date,1,7)=?",(month,)).fetchone()['v'] or 0)
        parts_rev=float(conn.execute("SELECT COALESCE(SUM(sale_price),0) v FROM part_sales WHERE substr(sale_date,1,7)=?",(month,)).fetchone()['v'] or 0)
        result["kind"]="money"; result["summary"]=f"Recorded revenue this month is {_copilot_money(vehicle_rev+parts_rev)}."
        result["items"]=[{"title":"Vehicle sales","subtitle":month,"amount":_copilot_money(vehicle_rev),"url":url_for('invoice_centre')},{"title":"Parts sales","subtitle":month,"amount":_copilot_money(parts_rev),"url":url_for('parts_page')}]
        return result

    # Broad safe search fallback.
    like=f"%{q}%"
    vehicles=conn.execute("SELECT id,stock_no,year,make,model,status FROM vehicles WHERE stock_no LIKE ? OR make LIKE ? OR model LIKE ? OR vin LIKE ? OR registration LIKE ? LIMIT 12",(like,like,like,like,like)).fetchall()
    parts=conn.execute("SELECT id,part_number,part_name,make,model,quantity_on_hand,selling_price FROM parts WHERE part_number LIKE ? OR part_name LIKE ? OR make LIKE ? OR model LIKE ? OR manufacturer_part_no LIKE ? LIMIT 12",(like,like,like,like,like)).fetchall()
    result["kind"]="search"
    for r in vehicles:
        result["items"].append({"title":f"{r['stock_no']} — {r['year'] or ''} {r['make']} {r['model']}","subtitle":r['status'],"amount":"","url":url_for('vehicle_detail',vehicle_id=r['id'])})
    for r in parts:
        result["items"].append({"title":f"{r['part_number'] or ''} — {r['part_name']}","subtitle":f"{r['make'] or ''} {r['model'] or ''} · Qty {r['quantity_on_hand'] or 0}","amount":_copilot_money(r['selling_price']),"url":url_for('part_detail',part_id=r['id'])})
    result["summary"]=f"I found {len(result['items'])} BAM record(s) matching “{q}”." if result['items'] else "I couldn't match that request yet. Try asking about parts, stock age, photos, workshop jobs, reminders, invoices or profitable vehicles."
    result["tips"]=["Which vehicles have been sitting longer than 90 days?","Show vehicles missing photos","Show overdue workshop jobs","Which vehicles are most profitable?","Show draft invoices","Find Mini Cooper engine parts"]
    return result


@app.route("/copilot", methods=["GET", "POST"])
@login_required
def dealer_copilot():
    question=(request.form.get("question") if request.method=="POST" else request.args.get("q")) or ""
    conn=db()
    answer=dealer_copilot_answer(conn,question) if question.strip() else None
    conn.close()
    examples=[
        "Which vehicles have been sitting longer than 90 days?",
        "Show vehicles missing photos",
        "Show overdue workshop jobs",
        "Which vehicles are most profitable?",
        "Show draft invoices",
        "Find Mini Cooper engine parts",
    ]
    return render_template("dealer_copilot.html",question=question,answer=answer,examples=examples)


@app.route("/executive-dashboard")
@login_required
def executive_dashboard():
    conn=db(); ensure_smart_reminders(conn)
    today=date.today(); today_text=today.isoformat(); month=today.strftime('%Y-%m'); year=today.strftime('%Y')
    vehicle_today=float(conn.execute("SELECT COALESCE(SUM(sale_price_inc_gst),0) v FROM sales WHERE substr(sale_date,1,10)=?",(today_text,)).fetchone()['v'] or 0)
    parts_today=float(conn.execute("SELECT COALESCE(SUM(sale_price),0) v FROM part_sales WHERE substr(sale_date,1,10)=?",(today_text,)).fetchone()['v'] or 0)
    vehicle_month=float(conn.execute("SELECT COALESCE(SUM(sale_price_inc_gst),0) v FROM sales WHERE substr(sale_date,1,7)=?",(month,)).fetchone()['v'] or 0)
    parts_month=float(conn.execute("SELECT COALESCE(SUM(sale_price),0) v FROM part_sales WHERE substr(sale_date,1,7)=?",(month,)).fetchone()['v'] or 0)
    vehicle_year=float(conn.execute("SELECT COALESCE(SUM(sale_price_inc_gst),0) v FROM sales WHERE substr(sale_date,1,4)=?",(year,)).fetchone()['v'] or 0)
    parts_year=float(conn.execute("SELECT COALESCE(SUM(sale_price),0) v FROM part_sales WHERE substr(sale_date,1,4)=?",(year,)).fetchone()['v'] or 0)
    month_cost=float(conn.execute("SELECT COALESCE(SUM(cost_inc_gst),0) v FROM expenses WHERE substr(expense_date,1,7)=?",(month,)).fetchone()['v'] or 0)
    month_cost+=float(conn.execute("SELECT COALESCE(SUM(CASE WHEN actual_cost_inc_gst>0 THEN actual_cost_inc_gst ELSE estimated_cost END),0) v FROM job_cards WHERE substr(job_date,1,7)=?",(month,)).fetchone()['v'] or 0)
    month_cost+=float(conn.execute("SELECT COALESCE(SUM(cost_inc_gst),0) v FROM service_entries WHERE substr(service_date,1,7)=?",(month,)).fetchone()['v'] or 0)
    active=conn.execute("SELECT COUNT(*) c,COALESCE(SUM(purchase_price_inc_gst),0) value FROM vehicles WHERE status NOT IN ('Sold','BER')").fetchone()
    avg_age=float(conn.execute("SELECT COALESCE(AVG(julianday(?) - julianday(COALESCE(purchase_date,substr(created_at,1,10)))),0) v FROM vehicles WHERE status NOT IN ('Sold','BER')",(today_text,)).fetchone()['v'] or 0)
    parts_stock=conn.execute("SELECT COALESCE(SUM(quantity_on_hand),0) qty,COALESCE(SUM(quantity_on_hand*selling_price),0) retail FROM parts").fetchone()
    open_workshop=conn.execute("SELECT COUNT(*) c FROM job_cards WHERE COALESCE(status,'Open')!='Completed'").fetchone()['c']
    open_reminders=conn.execute("SELECT COUNT(*) c FROM reminders WHERE completed=0").fetchone()['c']
    draft_invoices=conn.execute("SELECT COUNT(*) c,COALESCE(SUM(sale_price_inc_gst),0) value FROM sales WHERE COALESCE(invoice_status,'Draft')!='Paid'").fetchone()
    unread_emails=conn.execute("SELECT COUNT(*) c FROM email_messages WHERE status='Unread'").fetchone()['c']
    donor_count=conn.execute("SELECT COUNT(*) c FROM vehicles WHERE COALESCE(vehicle_purpose,'Retail Sale')='Parts Vehicle' OR status='BER'").fetchone()['c']
    top_donors=conn.execute("""
      SELECT v.id,v.stock_no,v.year,v.make,v.model,
       COALESCE((SELECT SUM(ps.sale_price) FROM part_sales ps JOIN parts p ON p.id=ps.part_id WHERE p.vehicle_id=v.id),0)+COALESCE(v.shell_sale_price,0) revenue,
       COALESCE(v.purchase_price_inc_gst,0)+COALESCE((SELECT SUM(e.cost_inc_gst) FROM expenses e WHERE e.vehicle_id=v.id),0) invested
      FROM vehicles v WHERE COALESCE(v.vehicle_purpose,'Retail Sale')='Parts Vehicle' OR v.status='BER'
      ORDER BY (revenue-invested) DESC LIMIT 6
    """).fetchall()
    aged_stock=conn.execute("""
      SELECT id,stock_no,year,make,model,status,CAST(julianday(?) - julianday(COALESCE(purchase_date,substr(created_at,1,10))) AS INTEGER) days
      FROM vehicles WHERE status NOT IN ('Sold','BER') ORDER BY days DESC LIMIT 8
    """,(today_text,)).fetchall()
    monthly=conn.execute("""
      WITH months AS (
        SELECT substr(sale_date,1,7) m FROM sales WHERE sale_date IS NOT NULL
        UNION SELECT substr(sale_date,1,7) FROM part_sales WHERE sale_date IS NOT NULL
      )
      SELECT m,
        COALESCE((SELECT SUM(sale_price_inc_gst) FROM sales WHERE substr(sale_date,1,7)=m),0) vehicle_sales,
        COALESCE((SELECT SUM(sale_price) FROM part_sales WHERE substr(sale_date,1,7)=m),0) parts_sales
      FROM months WHERE m<>'' GROUP BY m ORDER BY m DESC LIMIT 12
    """).fetchall()
    monthly=list(reversed(monthly))
    conn.close()
    kpis={"today_revenue":vehicle_today+parts_today,"month_revenue":vehicle_month+parts_month,"year_revenue":vehicle_year+parts_year,"month_result":vehicle_month+parts_month-month_cost,
          "active_stock":active['c'],"stock_investment":active['value'],"avg_age":avg_age,"parts_qty":parts_stock['qty'],"parts_retail":parts_stock['retail'],"open_workshop":open_workshop,
          "open_reminders":open_reminders,"draft_invoices":draft_invoices['c'],"draft_invoice_value":draft_invoices['value'],"unread_emails":unread_emails,"donor_count":donor_count}
    return render_template("executive_dashboard.html",kpis=kpis,top_donors=top_donors,aged_stock=aged_stock,
                           month_labels=[r['m'] for r in monthly],month_values=[float(r['vehicle_sales'] or 0)+float(r['parts_sales'] or 0) for r in monthly])


# -----------------------------------------------------------------------------
# Version 24.1 - Professional Workshop Scheduler
# -----------------------------------------------------------------------------

def _parse_iso_day(value, fallback=None):
    try:
        return datetime.strptime(value or "", "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return fallback or date.today()


def _workshop_bays():
    raw = os.environ.get("BAM_WORKSHOP_BAYS", "Bay 1,Bay 2,Bay 3")
    return [item.strip() for item in raw.split(",") if item.strip()]


@app.route("/workshop-scheduler", methods=["GET", "POST"])
@login_required
def workshop_scheduler():
    if request.method == "POST":
        conn = db()
        try:
            booking_date = (request.form.get("booking_date") or "").strip()
            description = (request.form.get("description") or "").strip()
            if not booking_date:
                raise ValueError("Booking date is required.")
            _parse_iso_day(booking_date)
            if not description:
                raise ValueError("Enter the workshop work required.")
            vehicle_id = int(request.form.get("vehicle_id")) if request.form.get("vehicle_id") else None
            contact_id = int(request.form.get("contact_id")) if request.form.get("contact_id") else None
            quoted_hours = max(0.0, float(request.form.get("quoted_hours") or 0))
            labour_rate = max(0.0, float(request.form.get("labour_rate") or DEFAULT_LABOUR_RATE))
            cur = conn.execute("""
                INSERT INTO workshop_bookings(
                    vehicle_id,contact_id,booking_date,start_time,end_time,technician,bay,job_type,description,
                    status,quoted_hours,labour_rate,customer_name,customer_phone,notes,created_by,updated_at
                ) VALUES(?,?,?,?,?,?,?,?,?,'Booked',?,?,?,?,?,?,?,?)
            """, (
                vehicle_id, contact_id, booking_date, request.form.get("start_time") or None,
                request.form.get("end_time") or None, (request.form.get("technician") or "").strip() or None,
                (request.form.get("bay") or "").strip() or None, (request.form.get("job_type") or "General").strip(),
                description, quoted_hours, labour_rate, (request.form.get("customer_name") or "").strip() or None,
                (request.form.get("customer_phone") or "").strip() or None, (request.form.get("notes") or "").strip() or None,
                session.get("display_name"), datetime.now().isoformat(timespec="seconds")
            ))
            booking_id = cur.lastrowid
            conn.commit()
            log_action("Workshop booking created", "workshop_booking", booking_id, f"{booking_date} — {description}")
            flash("Workshop booking created.", "success")
            return redirect(url_for("workshop_booking_detail", booking_id=booking_id))
        except (ValueError, sqlite3.Error) as exc:
            conn.rollback(); flash(str(exc), "error")
        finally:
            conn.close()

    selected = _parse_iso_day(request.args.get("date"), date.today())
    week_start = selected - timedelta(days=selected.weekday())
    week_days = [week_start + timedelta(days=i) for i in range(7)]
    start_iso, end_iso = week_days[0].isoformat(), week_days[-1].isoformat()
    conn = db()
    rows = conn.execute("""
        SELECT b.*,v.stock_no,v.year,v.make,v.model,v.registration,c.name AS contact_name
        FROM workshop_bookings b
        LEFT JOIN vehicles v ON v.id=b.vehicle_id
        LEFT JOIN contacts c ON c.id=b.contact_id
        WHERE b.booking_date BETWEEN ? AND ?
        ORDER BY b.booking_date,COALESCE(b.start_time,'99:99'),b.id
    """, (start_iso, end_iso)).fetchall()
    vehicles = conn.execute("""
        SELECT id,stock_no,year,make,model,registration FROM vehicles
        WHERE status NOT IN ('Sold') ORDER BY stock_no
    """).fetchall()
    contacts = conn.execute("SELECT id,name,phone FROM contacts ORDER BY name").fetchall()
    users = conn.execute("SELECT display_name FROM users WHERE is_active=1 ORDER BY display_name").fetchall()
    today_iso = date.today().isoformat()
    today_stats = conn.execute("""
        SELECT COUNT(*) AS bookings,
               COALESCE(SUM(quoted_hours),0) AS quoted_hours,
               SUM(CASE WHEN status='Completed' THEN 1 ELSE 0 END) AS completed,
               SUM(CASE WHEN status='In Progress' THEN 1 ELSE 0 END) AS in_progress
        FROM workshop_bookings WHERE booking_date=?
    """, (today_iso,)).fetchone()
    open_timers = conn.execute("""
        SELECT t.*,b.description,b.technician,b.bay,v.stock_no,v.make,v.model
        FROM workshop_time_entries t
        JOIN workshop_bookings b ON b.id=t.booking_id
        LEFT JOIN vehicles v ON v.id=b.vehicle_id
        WHERE t.clock_out IS NULL ORDER BY t.clock_in
    """).fetchall()
    conn.close()
    by_day = {d.isoformat(): [] for d in week_days}
    for row in rows:
        by_day.setdefault(row["booking_date"], []).append(row)
    prev_week=(week_start-timedelta(days=7)).isoformat(); next_week=(week_start+timedelta(days=7)).isoformat()
    return render_template("workshop_scheduler.html", week_days=week_days, by_day=by_day, selected=selected,
                           prev_week=prev_week, next_week=next_week, vehicles=vehicles, contacts=contacts,
                           technicians=[r["display_name"] for r in users if r["display_name"]], bays=_workshop_bays(),
                           default_labour_rate=DEFAULT_LABOUR_RATE, today_stats=today_stats, open_timers=open_timers,
                           today_iso=today_iso)


@app.route("/workshop-bookings/<int:booking_id>", methods=["GET", "POST"])
@login_required
def workshop_booking_detail(booking_id):
    conn = db()
    booking = conn.execute("""
        SELECT b.*,v.stock_no,v.year,v.make,v.model,v.registration,v.vin,c.name AS contact_name
        FROM workshop_bookings b
        LEFT JOIN vehicles v ON v.id=b.vehicle_id
        LEFT JOIN contacts c ON c.id=b.contact_id
        WHERE b.id=?
    """, (booking_id,)).fetchone()
    if not booking:
        conn.close(); return "Workshop booking not found", 404
    if request.method == "POST":
        try:
            quoted_hours=max(0.0,float(request.form.get("quoted_hours") or 0))
            labour_rate=max(0.0,float(request.form.get("labour_rate") or DEFAULT_LABOUR_RATE))
            status=request.form.get("status") or "Booked"
            allowed={"Booked","Checked In","In Progress","Waiting Parts","Ready","Completed","Cancelled"}
            if status not in allowed: raise ValueError("Invalid workshop status.")
            conn.execute("""
                UPDATE workshop_bookings SET booking_date=?,start_time=?,end_time=?,technician=?,bay=?,job_type=?,description=?,
                    status=?,quoted_hours=?,labour_rate=?,customer_name=?,customer_phone=?,notes=?,updated_at=? WHERE id=?
            """, (request.form.get("booking_date") or booking["booking_date"],request.form.get("start_time") or None,
                  request.form.get("end_time") or None,(request.form.get("technician") or "").strip() or None,
                  (request.form.get("bay") or "").strip() or None,request.form.get("job_type") or "General",
                  (request.form.get("description") or "").strip() or booking["description"],status,quoted_hours,labour_rate,
                  (request.form.get("customer_name") or "").strip() or None,(request.form.get("customer_phone") or "").strip() or None,
                  (request.form.get("notes") or "").strip() or None,datetime.now().isoformat(timespec="seconds"),booking_id))
            conn.commit(); log_action("Workshop booking updated","workshop_booking",booking_id,status)
            flash("Workshop booking updated.","success")
            return redirect(url_for("workshop_booking_detail",booking_id=booking_id))
        except (ValueError,sqlite3.Error) as exc:
            conn.rollback(); flash(str(exc),"error")
    time_entries=conn.execute("SELECT * FROM workshop_time_entries WHERE booking_id=? ORDER BY clock_in DESC,id DESC",(booking_id,)).fetchall()
    total_minutes=sum(int(r["minutes"] or 0) for r in time_entries)
    open_timer=next((r for r in time_entries if not r["clock_out"]),None)
    conn.close()
    return render_template("workshop_booking.html",booking=booking,time_entries=time_entries,total_minutes=total_minutes,
                           open_timer=open_timer,bays=_workshop_bays(),default_labour_rate=DEFAULT_LABOUR_RATE)


@app.post("/workshop-bookings/<int:booking_id>/clock-in")
@login_required
def workshop_booking_clock_in(booking_id):
    conn=db()
    try:
        booking=conn.execute("SELECT * FROM workshop_bookings WHERE id=?",(booking_id,)).fetchone()
        if not booking: return "Workshop booking not found",404
        existing=conn.execute("SELECT id FROM workshop_time_entries WHERE booking_id=? AND clock_out IS NULL",(booking_id,)).fetchone()
        if existing: raise ValueError("This booking already has an active timer.")
        tech=(request.form.get("technician") or booking["technician"] or session.get("display_name") or "Technician").strip()
        now=datetime.now().isoformat(timespec="seconds")
        conn.execute("INSERT INTO workshop_time_entries(booking_id,technician,clock_in,notes,created_by) VALUES(?,?,?,?,?)",
                     (booking_id,tech,now,request.form.get("notes"),session.get("display_name")))
        conn.execute("UPDATE workshop_bookings SET technician=COALESCE(technician,?),status='In Progress',updated_at=? WHERE id=?",(tech,now,booking_id))
        conn.commit(); log_action("Workshop timer started","workshop_booking",booking_id,tech); flash(f"Clocked on: {tech}.","success")
    except (ValueError,sqlite3.Error) as exc:
        conn.rollback(); flash(str(exc),"error")
    finally: conn.close()
    return redirect(request.referrer or url_for("workshop_booking_detail",booking_id=booking_id))


@app.post("/workshop-bookings/<int:booking_id>/clock-out")
@login_required
def workshop_booking_clock_out(booking_id):
    conn=db()
    try:
        entry=conn.execute("SELECT * FROM workshop_time_entries WHERE booking_id=? AND clock_out IS NULL ORDER BY id DESC LIMIT 1",(booking_id,)).fetchone()
        if not entry: raise ValueError("There is no active timer for this booking.")
        out=datetime.now(); started=datetime.fromisoformat(entry["clock_in"])
        minutes=max(0,int(round((out-started).total_seconds()/60)))
        conn.execute("UPDATE workshop_time_entries SET clock_out=?,minutes=? WHERE id=?",(out.isoformat(timespec="seconds"),minutes,entry["id"]))
        if request.form.get("complete") == "1":
            conn.execute("UPDATE workshop_bookings SET status='Completed',updated_at=? WHERE id=?",(out.isoformat(timespec="seconds"),booking_id))
        conn.commit(); log_action("Workshop timer stopped","workshop_booking",booking_id,f"{minutes} minutes"); flash(f"Clocked off — {minutes} minutes recorded.","success")
    except (ValueError,sqlite3.Error) as exc:
        conn.rollback(); flash(str(exc),"error")
    finally: conn.close()
    return redirect(request.referrer or url_for("workshop_booking_detail",booking_id=booking_id))


@app.post("/workshop-bookings/<int:booking_id>/create-job-card")
@login_required
def workshop_booking_create_job(booking_id):
    conn=db()
    try:
        b=conn.execute("SELECT * FROM workshop_bookings WHERE id=?",(booking_id,)).fetchone()
        if not b: return "Workshop booking not found",404
        if not b["vehicle_id"]: raise ValueError("Link a vehicle to this booking before creating a job card.")
        if b["job_card_id"]:
            flash(f"This booking is already linked to job card #{b['job_card_id']}.","error")
            return redirect(url_for("job_card_edit",job_id=b["job_card_id"]))
        hours=float(b["quoted_hours"] or 0); rate=float(b["labour_rate"] or DEFAULT_LABOUR_RATE); estimate=round(hours*rate,2)
        cur=conn.execute("""
            INSERT INTO job_cards(vehicle_id,job_date,category,description,supplier,paid_by,estimated_cost,actual_cost_inc_gst,gst_amount,status,notes,labour_hours,labour_rate,labour_source)
            VALUES(?,?,?,?,?,?,?,?,?,'Open',?,?,?,?)
        """,(b["vehicle_id"],b["booking_date"],b["job_type"] or "Workshop",b["description"],"","Business",estimate,0,0,b["notes"],hours,rate,"Workshop Scheduler"))
        job_id=cur.lastrowid
        conn.execute("INSERT INTO job_card_history(job_card_id,old_status,new_status,changed_by,note) VALUES(?,?,?,?,?)",
                     (job_id,None,"Open",session.get("display_name"),f"Created from workshop booking #{booking_id}"))
        conn.execute("UPDATE workshop_bookings SET job_card_id=?,updated_at=? WHERE id=?",(job_id,datetime.now().isoformat(timespec="seconds"),booking_id))
        conn.commit(); log_action("Workshop booking converted to job card","workshop_booking",booking_id,f"Job #{job_id}")
        flash(f"Job card #{job_id} created.","success"); return redirect(url_for("job_card_edit",job_id=job_id))
    except (ValueError,sqlite3.Error) as exc:
        conn.rollback(); flash(str(exc),"error")
    finally: conn.close()
    return redirect(url_for("workshop_booking_detail",booking_id=booking_id))


@app.post("/workshop-bookings/<int:booking_id>/delete")
@login_required
def workshop_booking_delete(booking_id):
    conn=db()
    try:
        b=conn.execute("SELECT * FROM workshop_bookings WHERE id=?",(booking_id,)).fetchone()
        if not b: return "Workshop booking not found",404
        if b["job_card_id"]: raise ValueError("This booking is linked to a job card. Cancel it instead of deleting it.")
        conn.execute("DELETE FROM workshop_bookings WHERE id=?",(booking_id,)); conn.commit()
        log_action("Workshop booking deleted","workshop_booking",booking_id,b["description"]); flash("Workshop booking deleted.","success")
    except (ValueError,sqlite3.Error) as exc:
        conn.rollback(); flash(str(exc),"error")
    finally: conn.close()
    return redirect(url_for("workshop_scheduler"))


@app.get("/parts-intelligence")
@login_required
def parts_intelligence():
    conn = db()
    today = date.today()
    stale_cutoff = (today - timedelta(days=180)).isoformat()
    metrics = conn.execute("""
        SELECT COUNT(*) AS lines,
               COALESCE(SUM(quantity_on_hand),0) AS units,
               COALESCE(SUM(quantity_on_hand*unit_cost_inc_gst),0) AS cost_value,
               COALESCE(SUM(quantity_on_hand*selling_price),0) AS retail_value,
               SUM(CASE WHEN quantity_on_hand<=reorder_level AND status NOT IN ('Sold','Scrap') THEN 1 ELSE 0 END) AS low_stock,
               SUM(CASE WHEN COALESCE(storage_location,'')='' AND quantity_on_hand>0 THEN 1 ELSE 0 END) AS no_location,
               SUM(CASE WHEN status='Reserved' THEN 1 ELSE 0 END) AS reserved
        FROM parts
    """).fetchone()
    no_photos = conn.execute("""
        SELECT COUNT(*) AS c FROM parts p
        WHERE p.quantity_on_hand>0 AND NOT EXISTS(SELECT 1 FROM part_photos ph WHERE ph.part_id=p.id)
    """).fetchone()["c"]
    stale = conn.execute("""
        SELECT p.*, CAST(julianday(?) - julianday(COALESCE(NULLIF(p.date_added,''),NULLIF(p.updated_at,''),date('now'))) AS INTEGER) AS days_old
        FROM parts p
        WHERE p.quantity_on_hand>0 AND COALESCE(NULLIF(p.date_added,''),NULLIF(p.updated_at,''),'9999-12-31')<=?
        ORDER BY days_old DESC LIMIT 12
    """, (today.isoformat(), stale_cutoff)).fetchall()
    attention = conn.execute("""
        SELECT p.*,
          (SELECT COUNT(*) FROM part_photos ph WHERE ph.part_id=p.id) AS photo_count
        FROM parts p
        WHERE p.quantity_on_hand>0 AND (
          COALESCE(p.storage_location,'')='' OR COALESCE(p.manufacturer_part_no,'')='' OR
          COALESCE(p.barcode,'')='' OR NOT EXISTS(SELECT 1 FROM part_photos ph WHERE ph.part_id=p.id)
        )
        ORDER BY p.selling_price DESC,p.id DESC LIMIT 15
    """).fetchall()
    categories = conn.execute("""
        SELECT COALESCE(NULLIF(category,''),'Uncategorised') AS category, COUNT(*) AS lines,
               COALESCE(SUM(quantity_on_hand),0) AS units,
               COALESCE(SUM(quantity_on_hand*selling_price),0) AS retail
        FROM parts WHERE quantity_on_hand>0
        GROUP BY COALESCE(NULLIF(category,''),'Uncategorised') ORDER BY retail DESC LIMIT 12
    """).fetchall()
    recent_sales = conn.execute("""
        SELECT ps.*,p.part_number,p.part_name,p.vehicle_stock_no
        FROM part_sales ps JOIN parts p ON p.id=ps.part_id
        ORDER BY ps.sale_date DESC,ps.id DESC LIMIT 12
    """).fetchall()
    conn.close()
    gross_margin = float(metrics["retail_value"] or 0) - float(metrics["cost_value"] or 0)
    return render_template("parts_intelligence.html", metrics=metrics, no_photos=no_photos, stale=stale,
                           attention=attention, categories=categories, recent_sales=recent_sales, gross_margin=gross_margin)


@app.get("/parts/scan")
@login_required
def parts_scanner():
    return render_template("parts_scanner.html")


@app.get("/parts/lookup")
@login_required
def part_lookup():
    code = (request.args.get("code") or "").strip()
    if not code:
        flash("Scan or enter a part number, barcode or OEM number.", "error")
        return redirect(url_for("parts_scanner"))
    # QR labels contain a BAM part URL. Resolve the trailing numeric id safely.
    match = re.search(r"/parts/(\d+)(?:$|[/?#])", code)
    conn = db()
    if match:
        row = conn.execute("SELECT id FROM parts WHERE id=?", (int(match.group(1)),)).fetchone()
    else:
        row = conn.execute("""
            SELECT id FROM parts
            WHERE LOWER(COALESCE(part_number,''))=LOWER(?)
               OR LOWER(COALESCE(barcode,''))=LOWER(?)
               OR LOWER(COALESCE(manufacturer_part_no,''))=LOWER(?)
               OR LOWER(COALESCE(alternate_part_numbers,'')) LIKE LOWER(?)
            ORDER BY CASE WHEN LOWER(COALESCE(part_number,''))=LOWER(?) THEN 0 ELSE 1 END,id DESC LIMIT 1
        """, (code, code, code, f"%{code}%", code)).fetchone()
    conn.close()
    if not row:
        flash(f"No part found for {code}.", "error")
        return redirect(url_for("parts_scanner", code=code))
    return redirect(url_for("part_detail", part_id=row["id"]))


@app.get("/parts/<int:part_id>/qr.svg")
@login_required
def part_qr(part_id):
    conn = db()
    part = conn.execute("SELECT id,part_number FROM parts WHERE id=?", (part_id,)).fetchone()
    conn.close()
    if not part:
        return "Part not found", 404
    if qrcode is None:
        return Response("QR support is not installed. Deploy requirements.txt and restart the app.", status=503, mimetype="text/plain")
    target = url_for("part_detail", part_id=part_id, _external=True)
    factory = qrcode.image.svg.SvgPathImage
    image = qrcode.make(target, image_factory=factory, box_size=8, border=2)
    out = io.BytesIO()
    image.save(out)
    return Response(out.getvalue(), mimetype="image/svg+xml", headers={"Cache-Control":"no-store"})


@app.post("/parts/<int:part_id>/apply-price")
@login_required
def part_apply_suggested_price(part_id):
    conn = db()
    try:
        part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
        if not part:
            return "Part not found", 404
        suggestion = parts_price_suggestion(conn, part)
        price = float(request.form.get("price") or suggestion["suggested"] or 0)
        if price < 0:
            raise ValueError("Price cannot be negative.")
        conn.execute("UPDATE parts SET selling_price=?,updated_at=CURRENT_TIMESTAMP WHERE id=?", (price, part_id))
        conn.commit()
        log_action("Parts Intelligence price applied", "part", part_id, f"{part['part_number']} -> ${price:.2f}")
        flash(f"Selling price updated to ${price:,.2f}.", "success")
    except (ValueError, sqlite3.Error) as exc:
        conn.rollback()
        flash(str(exc), "error")
    finally:
        conn.close()
    return redirect(url_for("part_detail", part_id=part_id))


@app.post("/parts/<int:part_id>/move")
@login_required
def part_move_location(part_id):
    new_location = (request.form.get("storage_location") or "").strip()
    if not new_location:
        flash("Enter the new rack / shelf / bin location.", "error")
        return redirect(url_for("part_detail", part_id=part_id))
    conn = db()
    try:
        part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
        if not part:
            return "Part not found", 404
        old_location = part["storage_location"] or ""
        conn.execute("UPDATE parts SET storage_location=?,inventory_last_checked=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",
                     (new_location, date.today().isoformat(), part_id))
        conn.execute("INSERT INTO part_location_history(part_id,old_location,new_location,moved_by,notes) VALUES(?,?,?,?,?)",
                     (part_id, old_location, new_location, session.get("display_name") or session.get("username") or "User", request.form.get("notes")))
        conn.commit()
        log_action("Part moved", "part", part_id, f"{old_location or 'Unlocated'} -> {new_location}")
        flash(f"{part['part_number'] or part['part_name']} moved to {new_location}.", "success")
    except sqlite3.Error as exc:
        conn.rollback(); flash(str(exc), "error")
    finally:
        conn.close()
    return redirect(url_for("part_detail", part_id=part_id))


@app.post("/parts/<int:part_id>/stock-check")
@login_required
def part_stock_check(part_id):
    conn = db()
    try:
        part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
        if not part:
            return "Part not found", 404
        conn.execute("UPDATE parts SET inventory_last_checked=?,updated_at=CURRENT_TIMESTAMP WHERE id=?", (date.today().isoformat(), part_id))
        conn.commit()
        log_action("Part stock checked", "part", part_id, part["part_number"] or part["part_name"])
        flash("Stock check recorded.", "success")
    except sqlite3.Error as exc:
        conn.rollback(); flash(str(exc), "error")
    finally:
        conn.close()
    return redirect(url_for("part_detail", part_id=part_id))


@app.route("/parts/<int:part_id>/label")
@login_required
def part_label(part_id):
    conn = db()
    part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
    if not part:
        conn.close()
        return "Part not found", 404
    donor = conn.execute("SELECT * FROM vehicles WHERE id=?", (part["vehicle_id"],)).fetchone() if part["vehicle_id"] else None
    conn.close()
    return render_template("part_label.html", part=part, donor=donor)


@app.route("/uploads/<path:filename>")
@login_required
def uploaded_file(filename):
    return send_from_directory(UPLOAD_DIR, filename)

@app.template_filter("money")
def money(value):
    try:
        return f"${float(value or 0):,.2f}"
    except Exception:
        return "$0.00"



# ---------------- Version 24.4: CRM + Communications ----------------

def _contact_purchase_summary(conn, contact):
    email=(contact["email"] or "").strip()
    phone=(contact["phone"] or "").strip()
    vehicle_rows=[]; part_rows=[]
    if email or phone:
        clauses=[]; params=[]
        if email:
            clauses.append("LOWER(TRIM(COALESCE(s.buyer_email,'')))=LOWER(TRIM(?))")
            params.append(email)
        digits=''.join(ch for ch in phone if ch.isdigit())
        if digits:
            clauses.append("REPLACE(REPLACE(REPLACE(REPLACE(COALESCE(s.buyer_phone,''),' ',''),'-',''),'(',''),')','') LIKE ?")
            params.append('%'+digits[-8:]+'%')
        if clauses:
            vehicle_rows=conn.execute(f"""
                SELECT 'Vehicle' purchase_type,s.sale_date purchase_date,s.sale_price_inc_gst amount,s.invoice_number,
                       v.id vehicle_id,v.stock_no item_no,(COALESCE(v.year,'') || ' ' || v.make || ' ' || v.model) item_name
                FROM sales s JOIN vehicles v ON v.id=s.vehicle_id
                WHERE {' OR '.join(clauses)} ORDER BY s.sale_date DESC
            """, params).fetchall()
        clauses=[]; params=[]
        if email:
            clauses.append("LOWER(TRIM(COALESCE(ps.customer_email,'')))=LOWER(TRIM(?))")
            params.append(email)
        if digits:
            clauses.append("REPLACE(REPLACE(REPLACE(REPLACE(COALESCE(ps.customer_phone,''),' ',''),'-',''),'(',''),')','') LIKE ?")
            params.append('%'+digits[-8:]+'%')
        if clauses:
            part_rows=conn.execute(f"""
                SELECT 'Part' purchase_type,ps.sale_date purchase_date,ps.sale_price amount,ps.invoice_number,
                       p.vehicle_id,p.part_number item_no,p.part_name item_name
                FROM part_sales ps JOIN parts p ON p.id=ps.part_id
                WHERE {' OR '.join(clauses)} ORDER BY ps.sale_date DESC
            """, params).fetchall()
    rows=[dict(r) for r in vehicle_rows]+[dict(r) for r in part_rows]
    rows.sort(key=lambda x: str(x.get('purchase_date') or ''), reverse=True)
    total=sum(float(r.get('amount') or 0) for r in rows)
    return rows,total


@app.route("/crm")
@login_required
def crm_centre():
    conn=db(); q=request.args.get("q","").strip(); status=request.args.get("status","").strip(); today=date.today().isoformat()
    where=[]; params=[]
    if q:
        pat=f"%{q}%"; where.append("(name LIKE ? OR phone LIKE ? OR email LIKE ? OR company LIKE ? OR tags LIKE ?)"); params += [pat]*5
    if status:
        where.append("crm_status=?"); params.append(status)
    sql="SELECT * FROM contacts" + (" WHERE "+" AND ".join(where) if where else "") + " ORDER BY COALESCE(next_follow_up_date,'9999-12-31'), name"
    contacts=[dict(r) for r in conn.execute(sql,params).fetchall()]
    for c in contacts:
        tx,total=_contact_purchase_summary(conn,c); c['purchase_count']=len(tx); c['total_spend']=total
        c['email_count']=conn.execute("SELECT COUNT(*) c FROM email_messages WHERE contact_id=?",(c['id'],)).fetchone()['c']
    metrics={
        'total':conn.execute("SELECT COUNT(*) c FROM contacts").fetchone()['c'],
        'leads':conn.execute("SELECT COUNT(*) c FROM contacts WHERE crm_status='Lead'").fetchone()['c'],
        'due':conn.execute("SELECT COUNT(*) c FROM contacts WHERE COALESCE(next_follow_up_date,'')<>'' AND next_follow_up_date<=? AND crm_status<>'Inactive'",(today,)).fetchone()['c'],
        'vip':conn.execute("SELECT COUNT(*) c FROM contacts WHERE crm_status='VIP'").fetchone()['c'],
    }
    due=conn.execute("SELECT * FROM contacts WHERE COALESCE(next_follow_up_date,'')<>'' AND next_follow_up_date<=? AND crm_status<>'Inactive' ORDER BY next_follow_up_date LIMIT 12",(today,)).fetchall()
    conn.close(); return render_template('crm_centre.html',contacts=contacts,metrics=metrics,due=due,q=q,status_filter=status,today=today)


@app.route("/crm/contact/<int:contact_id>", methods=["GET","POST"])
@login_required
def crm_contact_detail(contact_id):
    conn=db(); contact=conn.execute("SELECT * FROM contacts WHERE id=?",(contact_id,)).fetchone()
    if not contact: conn.close(); return "Contact not found",404
    if request.method=='POST':
        action=request.form.get('action','activity')
        try:
            if action=='profile':
                conn.execute("""UPDATE contacts SET contact_type=?,name=?,company=?,phone=?,email=?,address=?,licence_no=?,crm_status=?,source=?,tags=?,preferred_contact=?,next_follow_up_date=?,notes=?,updated_at=CURRENT_TIMESTAMP WHERE id=?""",(
                    request.form.get('contact_type'),request.form.get('name'),request.form.get('company'),request.form.get('phone'),request.form.get('email'),request.form.get('address'),request.form.get('licence_no'),request.form.get('crm_status') or 'Active',request.form.get('source'),request.form.get('tags'),request.form.get('preferred_contact'),request.form.get('next_follow_up_date'),request.form.get('notes'),contact_id))
                flash('Customer profile updated.','success')
            else:
                act_date=request.form.get('activity_date') or date.today().isoformat(); follow=request.form.get('follow_up_date') or None
                conn.execute("""INSERT INTO crm_activities(contact_id,vehicle_id,activity_date,activity_type,subject,notes,outcome,follow_up_date,created_by) VALUES(?,?,?,?,?,?,?,?,?)""",(
                    contact_id,int(request.form.get('vehicle_id')) if request.form.get('vehicle_id') else None,act_date,request.form.get('activity_type') or 'Note',request.form.get('subject') or 'CRM activity',request.form.get('activity_notes'),request.form.get('outcome'),follow,session.get('display_name')))
                conn.execute("UPDATE contacts SET last_contact_date=?,next_follow_up_date=COALESCE(?,next_follow_up_date),updated_at=CURRENT_TIMESTAMP WHERE id=?",(act_date,follow,contact_id))
                if follow and request.form.get('create_reminder')=='1':
                    conn.execute("INSERT INTO reminders(vehicle_id,reminder_date,reminder_type,title,notes,completed) VALUES(NULL,?,?,?,?,0)",(follow,'Customer Follow Up',f"Follow up {contact['name']}",request.form.get('subject') or 'CRM activity'))
                flash('CRM activity saved.','success')
            conn.commit()
        except sqlite3.Error as exc:
            conn.rollback(); flash(f'Could not save CRM update: {exc}','error')
        finally: conn.close()
        return redirect(url_for('crm_contact_detail',contact_id=contact_id))
    cdict=dict(contact); purchases,total=_contact_purchase_summary(conn,cdict)
    emails=conn.execute("SELECT * FROM email_messages WHERE contact_id=? ORDER BY COALESCE(message_date,created_at) DESC LIMIT 30",(contact_id,)).fetchall()
    activities=conn.execute("SELECT a.*,v.stock_no FROM crm_activities a LEFT JOIN vehicles v ON v.id=a.vehicle_id WHERE a.contact_id=? ORDER BY a.activity_date DESC,a.id DESC",(contact_id,)).fetchall()
    vehicles=conn.execute("SELECT id,stock_no,year,make,model FROM vehicles ORDER BY stock_no DESC").fetchall()
    conn.close(); return render_template('crm_contact_detail.html',contact=cdict,purchases=purchases,total_spend=total,emails=emails,activities=activities,vehicles=vehicles,today=date.today().isoformat())


@app.route("/email-centre/<int:message_id>/crm-followup", methods=["POST"])
@login_required
def email_to_crm_followup(message_id):
    conn=db(); msg=conn.execute("SELECT * FROM email_messages WHERE id=?",(message_id,)).fetchone()
    if not msg: conn.close(); return "Email not found",404
    contact_id=msg['contact_id']
    if not contact_id:
        name=(msg['sender'] if msg['direction']=='Incoming' else msg['recipient']) or 'Email Contact'
        email=name if '@' in name else ''
        cur=conn.execute("INSERT INTO contacts(contact_type,name,email,crm_status,source) VALUES(?,?,?,?,?)",('Customer',name,email,'Lead','Email Centre'))
        contact_id=cur.lastrowid; conn.execute("UPDATE email_messages SET contact_id=? WHERE id=?",(contact_id,message_id))
    follow=request.form.get('follow_up_date') or (date.today()+timedelta(days=1)).isoformat()
    conn.execute("""INSERT INTO crm_activities(contact_id,vehicle_id,email_message_id,activity_date,activity_type,subject,notes,outcome,follow_up_date,created_by) VALUES(?,?,?,?,?,?,?,?,?,?)""",(
        contact_id,msg['vehicle_id'],message_id,date.today().isoformat(),'Email Follow Up',msg['subject'],msg['body_excerpt'],'Pending',follow,session.get('display_name')))
    conn.execute("UPDATE contacts SET next_follow_up_date=?,last_contact_date=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(follow,date.today().isoformat(),contact_id))
    conn.execute("UPDATE email_messages SET status='Follow Up' WHERE id=?",(message_id,))
    conn.commit(); conn.close(); flash('Email converted to CRM follow-up.','success')
    return redirect(url_for('crm_contact_detail',contact_id=contact_id))



# -----------------------------------------------------------------------------
# Version 24.5 - Mobile Operations
# -----------------------------------------------------------------------------

@app.get("/mobile")
@login_required
def mobile_operations():
    today_iso = date.today().isoformat()
    conn = db()
    metrics = conn.execute("""
        SELECT
          (SELECT COUNT(*) FROM vehicles WHERE status NOT IN ('Sold')) AS vehicles,
          (SELECT COUNT(*) FROM parts WHERE quantity_on_hand>0) AS parts_in_stock,
          (SELECT COUNT(*) FROM workshop_bookings WHERE booking_date=? AND status!='Cancelled') AS workshop_today,
          (SELECT COUNT(*) FROM workshop_time_entries WHERE clock_out IS NULL) AS live_timers,
          (SELECT COUNT(*) FROM reminders WHERE completed=0 AND due_date<=?) AS reminders_due
    """, (today_iso, today_iso)).fetchone()
    bookings = conn.execute("""
        SELECT b.*,v.stock_no,v.year,v.make,v.model
        FROM workshop_bookings b
        LEFT JOIN vehicles v ON v.id=b.vehicle_id
        WHERE b.booking_date=? AND b.status!='Cancelled'
        ORDER BY COALESCE(b.start_time,'99:99'),b.id LIMIT 10
    """, (today_iso,)).fetchall()
    open_timers = conn.execute("""
        SELECT t.*,b.description,b.bay,v.stock_no,v.make,v.model
        FROM workshop_time_entries t
        JOIN workshop_bookings b ON b.id=t.booking_id
        LEFT JOIN vehicles v ON v.id=b.vehicle_id
        WHERE t.clock_out IS NULL ORDER BY t.clock_in
    """).fetchall()
    recent_vehicles = conn.execute("""
        SELECT id,stock_no,year,make,model,registration,vin,status
        FROM vehicles ORDER BY id DESC LIMIT 6
    """).fetchall()
    conn.close()
    return render_template("mobile_dashboard.html", metrics=metrics, bookings=bookings,
                           open_timers=open_timers, recent_vehicles=recent_vehicles, today_iso=today_iso)


@app.get("/mobile/vehicles")
@login_required
def mobile_vehicle_lookup():
    code = (request.args.get("code") or "").strip()
    rows = []
    if code:
        like = f"%{code}%"
        conn = db()
        rows = conn.execute("""
            SELECT id,stock_no,year,make,model,variant,vin,registration,status
            FROM vehicles
            WHERE LOWER(COALESCE(stock_no,''))=LOWER(?)
               OR LOWER(COALESCE(vin,''))=LOWER(?)
               OR LOWER(COALESCE(registration,''))=LOWER(?)
               OR stock_no LIKE ? OR vin LIKE ? OR registration LIKE ?
               OR make LIKE ? OR model LIKE ?
            ORDER BY CASE
                WHEN LOWER(COALESCE(stock_no,''))=LOWER(?) THEN 0
                WHEN LOWER(COALESCE(vin,''))=LOWER(?) THEN 0
                WHEN LOWER(COALESCE(registration,''))=LOWER(?) THEN 0
                ELSE 1 END,id DESC LIMIT 20
        """, (code,code,code,like,like,like,like,like,code,code,code)).fetchall()
        conn.close()
        if len(rows) == 1:
            return redirect(url_for("mobile_vehicle", vehicle_id=rows[0]["id"]))
    return render_template("mobile_vehicle_lookup.html", rows=rows, code=code)


@app.route("/mobile/vehicles/<int:vehicle_id>", methods=["GET", "POST"])
@login_required
def mobile_vehicle(vehicle_id):
    conn = db()
    vehicle = conn.execute("SELECT * FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    if not vehicle:
        conn.close(); return "Vehicle not found", 404
    if request.method == "POST":
        files = request.files.getlist("photos")
        caption = (request.form.get("caption") or "Mobile capture").strip() or "Mobile capture"
        added = 0
        try:
            for file in files:
                filename = save_upload(file)
                if filename:
                    conn.execute("INSERT INTO vehicle_photos(vehicle_id,filename,caption) VALUES(?,?,?)",
                                 (vehicle_id,filename,caption))
                    added += 1
            conn.commit()
            log_action("Mobile vehicle photo capture", "vehicle", vehicle_id, f"{added} photo(s)")
            flash(f"{added} vehicle photo(s) uploaded." if added else "No valid image selected.",
                  "success" if added else "error")
        except (ValueError, sqlite3.Error) as exc:
            conn.rollback(); flash(str(exc), "error")
        finally:
            conn.close()
        return redirect(url_for("mobile_vehicle", vehicle_id=vehicle_id))
    photos = conn.execute("SELECT * FROM vehicle_photos WHERE vehicle_id=? ORDER BY id DESC LIMIT 8", (vehicle_id,)).fetchall()
    parts_count = conn.execute("SELECT COUNT(*) AS c FROM parts WHERE vehicle_stock_no=?", (vehicle["stock_no"],)).fetchone()["c"]
    open_jobs = conn.execute("SELECT COUNT(*) AS c FROM job_cards WHERE vehicle_id=? AND COALESCE(status,'')!='Completed'", (vehicle_id,)).fetchone()["c"]
    conn.close()
    return render_template("mobile_vehicle.html", vehicle=vehicle, photos=photos, parts_count=parts_count, open_jobs=open_jobs)


@app.get("/mobile/parts")
@login_required
def mobile_part_lookup():
    code = (request.args.get("code") or "").strip()
    rows = []
    if code:
        match = re.search(r"/parts/(\d+)(?:$|[/?#])", code)
        conn = db()
        if match:
            row = conn.execute("SELECT id FROM parts WHERE id=?", (int(match.group(1)),)).fetchone()
            rows = [row] if row else []
        else:
            like = f"%{code}%"
            rows = conn.execute("""
                SELECT id,part_number,part_name,manufacturer_part_no,barcode,vehicle_stock_no,location,status,quantity_on_hand
                FROM parts
                WHERE LOWER(COALESCE(part_number,''))=LOWER(?)
                   OR LOWER(COALESCE(barcode,''))=LOWER(?)
                   OR LOWER(COALESCE(manufacturer_part_no,''))=LOWER(?)
                   OR part_number LIKE ? OR barcode LIKE ? OR manufacturer_part_no LIKE ?
                   OR part_name LIKE ? OR vehicle_stock_no LIKE ? OR location LIKE ?
                ORDER BY CASE WHEN LOWER(COALESCE(part_number,''))=LOWER(?) THEN 0 ELSE 1 END,id DESC LIMIT 20
            """, (code,code,code,like,like,like,like,like,like,code)).fetchall()
        conn.close()
        if len(rows) == 1:
            return redirect(url_for("mobile_part", part_id=rows[0]["id"]))
    return render_template("mobile_part_lookup.html", rows=rows, code=code)


@app.route("/mobile/parts/<int:part_id>", methods=["GET", "POST"])
@login_required
def mobile_part(part_id):
    conn = db()
    part = conn.execute("SELECT * FROM parts WHERE id=?", (part_id,)).fetchone()
    if not part:
        conn.close(); return "Part not found", 404
    if request.method == "POST":
        action = request.form.get("action") or "photo"
        try:
            if action == "photo":
                files = request.files.getlist("photos")
                added = 0
                for file in files:
                    filename = save_upload(file)
                    if filename:
                        featured = 1 if not conn.execute("SELECT id FROM part_photos WHERE part_id=? LIMIT 1", (part_id,)).fetchone() else 0
                        conn.execute("INSERT INTO part_photos(part_id,filename,caption,is_featured) VALUES(?,?,?,?)",
                                     (part_id,filename,"Mobile capture",featured))
                        added += 1
                conn.commit(); log_action("Mobile part photo capture","part",part_id,f"{added} photo(s)")
                flash(f"{added} part photo(s) uploaded." if added else "No valid image selected.", "success" if added else "error")
            elif action == "move":
                new_location = (request.form.get("location") or "").strip()
                if not new_location:
                    raise ValueError("Enter a rack, shelf or bin location.")
                old_location = part["location"] or ""
                conn.execute("UPDATE parts SET location=? WHERE id=?", (new_location,part_id))
                # The location history table was introduced in v24.3.
                conn.execute("INSERT INTO part_location_history(part_id,old_location,new_location,moved_by,notes) VALUES(?,?,?,?,?)",
                             (part_id,old_location,new_location,session.get("display_name"),"Mobile move"))
                conn.commit(); log_action("Mobile part moved","part",part_id,f"{old_location} -> {new_location}")
                flash(f"Part moved to {new_location}.","success")
            elif action == "stockcheck":
                conn.execute("UPDATE parts SET inventory_last_checked=CURRENT_TIMESTAMP WHERE id=?", (part_id,))
                conn.commit(); log_action("Mobile stock check","part",part_id,part["part_number"] or part["part_name"]); flash("Stock check recorded.","success")
            else:
                raise ValueError("Unknown mobile action.")
        except (ValueError,sqlite3.Error) as exc:
            conn.rollback(); flash(str(exc),"error")
        finally:
            conn.close()
        return redirect(url_for("mobile_part",part_id=part_id))
    photos = conn.execute("SELECT * FROM part_photos WHERE part_id=? ORDER BY is_featured DESC,id DESC LIMIT 6", (part_id,)).fetchall()
    conn.close()
    return render_template("mobile_part.html", part=part, photos=photos)


@app.get("/mobile/workshop")
@login_required
def mobile_workshop():
    today_iso = date.today().isoformat()
    conn = db()
    bookings = conn.execute("""
        SELECT b.*,v.stock_no,v.year,v.make,v.model,v.registration,
               (SELECT COUNT(*) FROM workshop_time_entries t WHERE t.booking_id=b.id AND t.clock_out IS NULL) AS timer_active
        FROM workshop_bookings b
        LEFT JOIN vehicles v ON v.id=b.vehicle_id
        WHERE b.booking_date=? AND b.status!='Cancelled'
        ORDER BY COALESCE(b.start_time,'99:99'),b.id
    """, (today_iso,)).fetchall()
    conn.close()
    return render_template("mobile_workshop.html", bookings=bookings, today_iso=today_iso)


# -----------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# BAM Buying Watch listing importer - Version 25.8
# Clean multi-site importer. Public listing pages are read with a browser-like
# request, parsed from metadata / JSON-LD / visible page text, and normalised
# into the same BAM Buying Watch fields. Sites that require login (especially
# Facebook Marketplace) fall back to pasted listing text instead of crashing.
# ---------------------------------------------------------------------------

def _strip_html(value):
    value = str(value or "")
    # Keep natural breaks between common block elements before stripping tags.
    value = re.sub(r"<(?:br|/p|/div|/li|/tr|/h[1-6])\b[^>]*>", "\n", value, flags=re.I)
    value = re.sub(r"<script\b[^>]*>.*?</script>", " ", value, flags=re.I | re.S)
    value = re.sub(r"<style\b[^>]*>.*?</style>", " ", value, flags=re.I | re.S)
    value = re.sub(r"<[^>]+>", " ", value)
    value = html.unescape(value)
    return re.sub(r"\s+", " ", value).strip()


def _meta_content(page_html, prop):
    patterns = [
        rf'<meta[^>]+(?:property|name)\s*=\s*["\']{re.escape(prop)}["\'][^>]+content\s*=\s*["\']([^"\']*)["\']',
        rf'<meta[^>]+content\s*=\s*["\']([^"\']*)["\'][^>]+(?:property|name)\s*=\s*["\']{re.escape(prop)}["\']',
    ]
    for pat in patterns:
        match = re.search(pat, page_html or "", flags=re.I | re.S)
        if match:
            return _strip_html(match.group(1))
    return ""


def _detect_listing_source(url):
    host = (urllib.parse.urlparse(url or "").netloc or "").lower()
    if "facebook.com" in host or "fb.com" in host:
        return "Facebook Marketplace"
    if "gumtree.com" in host:
        return "Gumtree"
    if "carsales.com" in host:
        return "Carsales"
    auction_words = ("manheim", "pickles", "grays", "lloyds", "slattery", "auction")
    if any(word in host for word in auction_words):
        return "Auction"
    return "Other"


def _listing_site_name(url):
    host = (urllib.parse.urlparse(url or "").netloc or "").lower()
    if "grays.com" in host:
        return "Grays"
    if "pickles.com" in host:
        return "Pickles"
    if "manheim.com" in host:
        return "Manheim"
    if "lloyds" in host:
        return "Lloyds Auctions"
    if "slattery" in host:
        return "Slattery Auctions"
    if "carsales.com" in host:
        return "Carsales"
    if "gumtree.com" in host:
        return "Gumtree"
    if "facebook.com" in host or "fb.com" in host:
        return "Facebook Marketplace"
    return ""


def _validate_public_http_url(url):
    """Reject local/private addresses before BAM makes a server-side request."""
    if not re.match(r"^https?://", url or "", flags=re.I):
        raise ValueError("Please paste a full http:// or https:// listing link.")
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in ("http", "https") or not parsed.hostname:
        raise ValueError("Only normal public web links can be imported.")
    host = parsed.hostname.lower()
    if host in {"localhost", "localhost.localdomain"}:
        raise ValueError("Local/private web addresses cannot be imported.")
    try:
        infos = socket.getaddrinfo(host, parsed.port or (443 if parsed.scheme == "https" else 80), type=socket.SOCK_STREAM)
        for info in infos:
            addr = ipaddress.ip_address(info[4][0])
            if addr.is_private or addr.is_loopback or addr.is_link_local or addr.is_multicast or addr.is_reserved:
                raise ValueError("Local/private web addresses cannot be imported.")
    except socket.gaierror as exc:
        raise ValueError(f"BAM could not find that website: {exc}")
    return parsed


def _first_match(text, patterns, flags=re.I):
    for pattern in patterns:
        match = re.search(pattern, text or "", flags=flags)
        if match:
            return match.group(1).strip()
    return ""


def _number(value, integer=False):
    if value is None:
        return None
    cleaned = re.sub(r"[^0-9.\-]", "", str(value).replace(",", ""))
    if not cleaned or cleaned in {"-", ".", "-."}:
        return None
    try:
        number = float(cleaned)
        return int(number) if integer else number
    except ValueError:
        return None


def _extract_listing_details(raw_text, url="", title="", description=""):
    clean = _strip_html(raw_text)
    source = _detect_listing_source(url)
    details = {"listing_source": source}
    source_text = clean[:180000]
    lower = source_text.lower()

    if url:
        details["listing_url"] = url
        if source == "Auction":
            details["auction_url"] = url
            site_name = _listing_site_name(url)
            if site_name:
                details["auction_name"] = site_name

    # Year.
    value = _first_match(source_text, [r"\b((?:19|20)\d{2})\b"])
    if value:
        details["year"] = int(value)

    # Prices. Labelled prices beat generic dollar amounts.
    asking = _first_match(source_text, [
        r"(?:asking\s+price|advertised\s+price|price)\s*[:\-]?\s*\$\s*([\d,]+(?:\.\d{1,2})?)",
    ])
    current_bid = _first_match(source_text, [
        r"current\s+bid(?:\s*\([^)]*\))?\s*[:\-]?\s*\$\s*([\d,]+(?:\.\d{1,2})?)",
        r"highest\s+bid\s*[:\-]?\s*\$\s*([\d,]+(?:\.\d{1,2})?)",
    ])
    if source == "Auction":
        val = _number(current_bid)
        if val is not None:
            details["current_bid"] = val
    else:
        val = _number(asking)
        if val is None:
            generic_price = _first_match(source_text, [r"\$\s*([\d,]{3,}(?:\.\d{1,2})?)"])
            val = _number(generic_price)
        if val is not None:
            details["asking_price"] = val

    # Kilometres / odometer.
    odometer = _first_match(source_text, [
        r"(?:indicated\s+odometer\s+reading|odometer\s+reading|odometer|kilometres|kilometers)\s*[:\-]?\s*([\d,]+)\s*(?:km|kms)?",
        r"([\d,]{2,})\s*(?:km|kms|kilometres|kilometers)\b",
    ])
    val = _number(odometer, integer=True)
    if val is not None and val < 10_000_000:
        details["odometer_km"] = val

    # VIN / chassis and registration.
    vin = _first_match(source_text, [
        r"(?:vin\s*/\s*chassis|vin|chassis(?:\s+number)?)\s*[:#\-]?\s*([A-HJ-NPR-Z0-9]{17})",
        r"\b([A-HJ-NPR-Z0-9]{17})\b",
    ])
    if vin:
        details["vin"] = vin.upper()
    rego = _first_match(source_text, [
        r"(?:registration\s+no\.?|registration|rego)\s*[:#\-]?\s*([A-Z0-9\-]{3,10})",
    ])
    if rego:
        details["registration"] = rego.upper()

    # Registration sale/status. Grays commonly publishes values such as
    # "Sold Registered, Sold on Consignment" or "Sold Unregistered".
    registration_status = _first_match(source_text, [
        r"(?:registration\s+status|rego\s+status)\s*[:\-]?\s*(.+?)(?=\s+(?:engine\s+capacity|engine\s+size|fuel\s+type|drive\s+type|transmission|indicated\s+odometer|odometer|exterior\s+colour|exterior\s+color|general\s+condition|features)\s*[:\-]?|$)",
    ])
    if registration_status:
        registration_status = re.sub(r"\s+", " ", registration_status).strip(" ,.-")
        if len(registration_status) <= 100:
            details["registration_status"] = registration_status

    # Reserve status, body type and seating capacity.
    if re.search(r"\bno\s+reserve\b", source_text, flags=re.I):
        details["reserve_status"] = "No Reserve"
    elif re.search(r"\breserve(?:\s+price)?(?:\s+applies|\s+met|\s+not\s+met|\s+price)?\b", source_text, flags=re.I):
        details["reserve_status"] = "Reserve"

    body_type = _first_match(source_text, [
        r"(?:body\s+type|body\s+style)\s*[:\-]?\s*([A-Za-z][A-Za-z /-]{1,30}?)(?=\s+(?:no\.?\s+of\s+seats|seats|build\s+date|compliance|vin|registration)\s*[:\-]?|$)",
    ])
    if body_type:
        details["body_type"] = re.sub(r"\s+", " ", body_type).strip(" ,.-").title()
    seats = _first_match(source_text, [
        r"(?:no\.?\s+of\s+seats|number\s+of\s+seats|seats)\s*[:\-]?\s*(\d{1,2})",
        r"\b(\d{1,2})\s*seat\b",
    ])
    if seats:
        try:
            seat_count = int(seats)
            if 1 <= seat_count <= 99:
                details["seat_count"] = seat_count
        except ValueError:
            pass

    # Engine details.
    engine_cc = _first_match(source_text, [
        r"(?:engine\s+capacity|engine\s+size|capacity)\s*[:\-]?\s*([\d,]{3,5})\s*cc\b",
        r"\b([\d,]{3,5})\s*cc\b",
    ])
    cc = _number(engine_cc, integer=True)
    if cc:
        details["engine_cc"] = cc
        litres = round(cc / 1000.0, 1)
        details.setdefault("engine_size", f"{litres:.1f}L")
    litre = _first_match(source_text, [r"\b(\d(?:\.\d)?)\s*(?:l|litre|liter)\b"])
    if litre and not details.get("engine_size"):
        details["engine_size"] = f"{litre}L"

    cylinders = _first_match(source_text, [
        r"(?:engine\s+cylinders|cylinders|cylinder)\s*[:\-]?\s*(\d{1,2})",
        r"\b(\d{1,2})\s*cyl(?:inder)?s?\b",
        r"\b(\d{1,2})cyl\b",
    ])
    if cylinders:
        details["engine_cylinders"] = cylinders

    # Transmission. Keep common BAM-compatible values but preserve Sports Automatic.
    transmission_checks = [
        ("Sports Automatic", r"\bsports\s+automatic\b|\bspts?\s+auto(?:matic)?\b"),
        ("DCT", r"\b(?:dct|dual[\s-]?clutch)\b"),
        ("CVT", r"\bcvt\b"),
        ("Automatic", r"\bautomatic\b"),
        ("Manual", r"\bmanual\b"),
    ]
    for label, pattern in transmission_checks:
        if re.search(pattern, source_text, flags=re.I):
            details["transmission"] = label
            break

    # Fuel and drive. Prefer a labelled fuel field. Auction/navigation pages can
    # contain unrelated words such as "hybrid", so broad page-wide matching is
    # deliberately avoided when a site does not publish a fuel label.
    labelled_fuel = _first_match(source_text, [
        r"(?:fuel\s+type|fuel)\s*[:\-]?\s*(plug[\s-]?in hybrid|phev|hybrid|electric|diesel|petrol|gasoline|lpg)\b",
    ])
    fuel_probe = labelled_fuel or " ".join(x for x in (title, description) if x)
    fuel_map = [
        ("Plug-in Hybrid", r"plug[\s-]?in hybrid|\bphev\b"),
        ("Hybrid", r"\bhybrid\b"),
        ("Electric", r"\belectric\b|\bev\b"),
        ("Diesel", r"\bdiesel\b"),
        ("Petrol", r"\bpetrol\b|\bgasoline\b"),
        ("LPG", r"\blpg\b"),
    ]
    for label, pattern in fuel_map:
        if re.search(pattern, fuel_probe, flags=re.I):
            details["fuel_type"] = label
            break
    drive_map = [
        ("4WD", r"\bfour\s+wheel\s+drive\b|\b4wd\b|\b4x4\b"),
        ("AWD", r"\ball[\s-]?wheel\s+drive\b|\bawd\b"),
        ("FWD", r"\bfront[\s-]?wheel\s+drive\b|\bfwd\b"),
        ("RWD", r"\brear[\s-]?wheel\s+drive\b|\brwd\b"),
        ("2WD", r"\b2wd\b"),
    ]
    for label, pattern in drive_map:
        if re.search(pattern, source_text, flags=re.I):
            details["drive_type"] = label
            break

    # Make/model using BAM catalogues.
    best_make = ""
    best_model = ""
    for asset_catalog in AUCTION_ASSET_MODEL_CATALOG.values():
        for make, models in asset_catalog.items():
            if re.search(rf"\b{re.escape(make.lower())}\b", lower):
                for model in sorted(models, key=len, reverse=True):
                    if re.search(rf"\b{re.escape(model.lower())}\b", lower):
                        best_make, best_model = make, model
                        break
                if best_make:
                    break
        if best_make:
            break
    if best_make:
        details["make"] = best_make
    if best_model:
        details["model"] = best_model

    # Facebook/Gumtree and older/classic vehicles are not always present in the
    # modern BAM catalogue. Recognise distinctive model names from the pasted ad
    # text rather than leaving Make / Model blank. These aliases are intentionally
    # conservative so BAM does not guess from ordinary description words.
    if not details.get("make") or not details.get("model"):
        classic_aliases = [
            ("Ford", "Model T", r"\b(?:ford\s+)?model\s+t\b"),
            ("Ford", "Model A", r"\bford\s+model\s+a\b"),
        ]
        for alias_make, alias_model, alias_pattern in classic_aliases:
            if re.search(alias_pattern, source_text, flags=re.I):
                details.setdefault("make", alias_make)
                details.setdefault("model", alias_model)
                best_make = details.get("make", alias_make)
                best_model = details.get("model", alias_model)
                break

    # For pasted Marketplace/Gumtree text, also inspect the first few non-empty
    # lines for a conventional vehicle heading such as "2012 Toyota Hilux SR5".
    # This supplements the catalogue matcher without overriding a confident match.
    if source in ("Facebook Marketplace", "Gumtree") and (not details.get("make") or not details.get("model")):
        heading_lines = [x.strip() for x in re.split(r"[\r\n]+", raw_text or "") if x.strip()][:8]
        heading_text = " ".join(heading_lines)[:1200]
        heading_lower = heading_text.lower()
        for asset_catalog in AUCTION_ASSET_MODEL_CATALOG.values():
            found = False
            for make, models in asset_catalog.items():
                if re.search(rf"\b{re.escape(make.lower())}\b", heading_lower):
                    for model in sorted(models, key=len, reverse=True):
                        if re.search(rf"\b{re.escape(model.lower())}\b", heading_lower):
                            details.setdefault("make", make)
                            details.setdefault("model", model)
                            best_make = details.get("make", make)
                            best_model = details.get("model", model)
                            found = True
                            break
                if found:
                    break
            if found:
                break

    # Variant: use BAM catalogue first, then capture text after detected model in title.
    if best_make and best_model:
        variants = VEHICLE_VARIANT_CATALOG.get(best_make, {}).get(best_model, [])
        for variant in sorted(variants, key=len, reverse=True):
            if re.search(rf"\b{re.escape(variant.lower())}\b", lower):
                details["variant"] = variant
                break
        title_text = _strip_html(title or "")
        year_prefix = rf"(?:19|20)\d{{2}}\s+" if details.get("year") else ""
        match = re.search(
            rf"\b{year_prefix}{re.escape(best_make)}\s+{re.escape(best_model)}\s+(.+?)(?=,|\s+-\s+|\s+(?:Sports\s+Automatic|Automatic|Manual|AWD|4WD|Petrol|Diesel|SUV)\b|$)",
            title_text,
            flags=re.I,
        )
        if match:
            candidate = re.sub(r"\s+", " ", match.group(1)).strip(" ,-")
            # Auction titles often include a badge + series (for example TX SY II),
            # which is more useful than a shorter catalogue-only badge such as TX.
            if 1 <= len(candidate) <= 40 and len(candidate) >= len(details.get("variant", "")):
                details["variant"] = candidate

    # Colour and interior.
    labelled_colour = _first_match(source_text, [
        r"(?:exterior\s+colour|exterior\s+color|colour|color)\s*[:\-]?\s*([A-Za-z][A-Za-z ]{1,24})",
    ])
    if labelled_colour:
        # Stop at the next likely label if stripped HTML ran labels together.
        labelled_colour = re.split(r"\b(?:general\s+condition|condition|interior|transmission|drive|fuel|registration|vin|engine|features)\b", labelled_colour, maxsplit=1, flags=re.I)[0].strip()
        details["colour"] = labelled_colour.title()
    else:
        colours = ("White", "Black", "Silver", "Grey", "Gray", "Blue", "Red", "Green", "Yellow", "Orange", "Brown", "Beige", "Gold", "Purple")
        for colour in colours:
            if re.search(rf"\b{colour.lower()}\b", lower):
                details["colour"] = "Grey" if colour == "Gray" else colour
                break
    interior = _first_match(source_text, [r"(?:interior|trim)\s*[:\-]?\s*(cloth|leather|vinyl|other)\b"])
    if interior:
        details["interior"] = interior.title()

    # Auction-specific labels used by Grays / Pickles / Manheim / Lloyds etc.
    lot = _first_match(source_text, [
        r"(?:lot\s+id|lot\s+no\.?|lot\s+number|lot)\s*[:#\-]?\s*([A-Z0-9\-]+)",
    ])
    if not lot and url:
        path_match = re.search(r"/lot/([A-Z0-9\-]+)", urllib.parse.urlparse(url).path, flags=re.I)
        if path_match:
            lot = path_match.group(1)
    if lot:
        details["lot_number"] = lot

    location = _first_match(source_text, [
        r"(?:auction\s+location|pickup\s+location|location)\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9 ,./&\-]{3,120})",
    ])
    if location:
        location = re.split(r"\b(?:category|current bid|lot|vin|registration|engine|odometer|colour|color|body type|features)\b", location, maxsplit=1, flags=re.I)[0].strip(" ,.-")
        if source == "Auction":
            details["auction_location"] = location
        else:
            details["seller_location"] = location

    # Conservative condition information. Do not invent Excellent/Good.
    condition_grade = _first_match(source_text, [r"(?:condition\s+grade|condition)\s*[:\-]?\s*(excellent|good|average|poor|damaged|ber|unknown)\b"])
    if condition_grade:
        details["condition_grade"] = condition_grade.title() if condition_grade.lower() != "ber" else "BER"

    note_parts = []
    if description:
        note_parts.append(_strip_html(description))
    elif source in ("Facebook Marketplace", "Gumtree") and raw_text:
        # Marketplace fallback imports are pasted as listing text rather than
        # fetched HTML. Preserve that seller description in BAM's
        # Condition / Inspection Notes so the original ad wording is kept with
        # the watch vehicle instead of being discarded after field extraction.
        marketplace_note = _strip_html(raw_text).strip()
        if marketplace_note and not marketplace_note.lower().startswith(("http://", "https://")):
            note_parts.append(marketplace_note)

    # Grays publishes two especially useful inspection sections. Preserve both
    # in BAM's Condition / Inspection Notes field so the buying decision has
    # the key/spare-key/service-history information and the assessor's damage
    # notes together in one place.
    if "grays.com" in (urllib.parse.urlparse(url or "").netloc or "").lower():
        general_condition = _first_match(source_text, [
            r"General\s+Condition\s+for\s+age\s+and\s+distance\s+travelled\s*:?\s*(.+?)(?=The\s+below\s+condition\s+assessment|Features\s*:|Motor\s+Dealer\s+Licence|$)",
        ], flags=re.I | re.S)
        condition_assessment = _first_match(source_text, [
            r"The\s+below\s+condition\s+assessment.*?(?:opinion\s*:?)\s*(.+?)(?=Features\s*:|Motor\s+Dealer\s+Licence|$)",
            r"The\s+below\s+condition\s+assessment\s+is\s+the\s+opinion\s+of\s+our\s+booking\s+staff.*?\s*(.+?)(?=Features\s*:|Motor\s+Dealer\s+Licence|$)",
        ], flags=re.I | re.S)

        if general_condition:
            # Restore readable line breaks between Grays' labelled checks.
            gc = re.sub(
                r"\s+(?=(?:Key|Spare\s+Key|Owners?\s+Manual|Service\s+History|Engine\s+Turns\s+Over)\s*:)",
                "\n",
                general_condition.strip(),
                flags=re.I,
            )
            note_parts.append("GENERAL CONDITION\n" + gc)

        if condition_assessment:
            ca = condition_assessment.strip()
            # Put numbered defects and the final free-text damage summary on
            # separate lines where Grays' HTML was flattened to one line.
            ca = re.sub(r"\s+(?=\d+\.\s+)", "\n", ca)
            ca = re.sub(
                r"\s+(?=Scratches\s+And\s+Dents|Scratches\s+and\s+Dents)",
                "\n",
                ca,
                flags=re.I,
            )
            note_parts.append("CONDITION ASSESSMENT\n" + ca)

    condition_text = _first_match(source_text, [
        r"(?:condition\s+details|inspection\s+notes|vehicle\s+condition)\s*[:\-]?\s*(.{20,1200}?)(?=\b(?:vin|registration|engine|odometer|location|current bid|lot)\b|$)",
    ], flags=re.I | re.S)
    if condition_text:
        note_parts.append(condition_text)

    # De-duplicate identical note blocks while preserving their order.
    unique_notes = []
    seen_notes = set()
    for part in note_parts:
        cleaned_part = (part or "").strip()
        key = re.sub(r"\s+", " ", cleaned_part).lower()
        if cleaned_part and key not in seen_notes:
            seen_notes.add(key)
            unique_notes.append(cleaned_part)

    note = "\n\n".join(unique_notes).strip()
    if note:
        details["condition_notes"] = note[:4000]

    # Extra Grays labels that are reliable on vehicle lot pages.
    if "grays.com" in (urllib.parse.urlparse(url or "").netloc or "").lower():
        sale_name = _first_match(source_text, [r"Part\s+of\s+Sale\s*[:\-]?\s*(.+?)(?=Warranty|Description|GST|Location|Lot\s+ID|$)"])
        if sale_name:
            details["auction_name"] = sale_name.strip(" -")[:120]

        # A completed lot may no longer expose a Current Bid. Keep its final
        # result in Sold Price instead of incorrectly treating it as a live bid.
        sold = _first_match(source_text, [
            r"(?:sold\s+for|sold\s+price|final\s+bid\s+price)\s*[:\-]?\s*\$\s*([\d,]+(?:\.\d{1,2})?)",
        ])
        sold_value = _number(sold)
        if sold_value is not None:
            details["sold_price"] = sold_value
            details["status"] = "Sold"

    # Version 25.18 - richer boat listing extraction for Grays and pasted Facebook/Gumtree text.
    boat_probe = " ".join((title or "", description or "", source_text[:50000]))
    if re.search(r"\b(boat|power\s*boat|runabout|bowrider|tinny|seaway|centre\s+console|outboard)\b", boat_probe, re.I):
        details["asset_type"] = "Boat"
        hin = _first_match(source_text, [r"(?:HIN|Hull\s+Identification\s+Number)\s*[:\-]?\s*([A-Z0-9-]{6,30})"], flags=re.I)
        if hin: details["vin"] = hin.upper()
        eng_make = _first_match(source_text, [r"Engine\s+Make\s*[:\-]?\s*([A-Za-z0-9 .&/-]{2,40})", r"Main\s+Engine\s+Details\s*[:\-]?\s*Engine\s+Make\s*[:\-]?\s*([A-Za-z0-9 .&/-]{2,40})"], flags=re.I)
        if eng_make: details["engine_make"] = eng_make.strip()
        hp = _first_match(source_text, [r"Horsepower\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*(?:HP|HP)?"], flags=re.I)
        if hp: details["horsepower"] = _number(hp)
        hours = _first_match(source_text, [r"Engine\s+Hours\s*[:\-]?\s*(?:Not\s+sure\s+)?(\d+(?:\.\d+)?)"], flags=re.I)
        if hours: details["engine_hours"] = _number(hours)
        fuel = _first_match(source_text, [r"Fuel\s+Type\s*[:\-]?\s*([A-Za-z0-9 -]{2,30})"], flags=re.I)
        if fuel: details["fuel_type"] = fuel.strip()
        length = _first_match(source_text, [r"Length\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*m(?:etres?)?"], flags=re.I)
        if length: details["length_m"] = _number(length)
        # v25.18.1 - Grays boat identity and mechanical details.
        # Boat titles such as "2016 Stacer Seaway 429" do not use the car
        # make/model catalogue, so capture the make/model directly from the title.
        clean_boat_title = _strip_html(title or "").strip()
        boat_title_match = re.search(r"\b(?:19|20)\d{2}\s+([A-Za-z][A-Za-z0-9&.'-]{1,30})\s+(.+?)(?=\s+-\s+|,|$)", clean_boat_title, re.I)
        if boat_title_match:
            boat_make = boat_title_match.group(1).strip()
            boat_model = re.sub(r"\s+", " ", boat_title_match.group(2)).strip(" -,")
            if boat_make:
                details["make"] = boat_make.title()
            if boat_model and len(boat_model) <= 80:
                details["model"] = boat_model

        beam = _first_match(source_text, [r"Beam\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*m(?:etres?)?"], flags=re.I)
        if beam:
            details["width_m"] = _number(beam)
        engine_type = _first_match(source_text, [r"Engine\s+Type\s*[:\-]?\s*([A-Za-z0-9 .&/-]{2,40})"], flags=re.I)
        engine_sn = _first_match(source_text, [r"Engine\s+(?:SN|Serial(?:\s+Number)?)\s*[:#\-]?\s*([A-Z0-9-]{4,40})"], flags=re.I)
        engine_turns = _first_match(source_text, [r"Engine\s+Turns\s+Over\s*[:\-]?\s*([A-Za-z]+)"], flags=re.I)
        rego_expiry = _first_match(source_text, [r"Rego\s+Expiry\s*[:\-]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})"], flags=re.I)
        depth = _first_match(source_text, [r"Depth\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*m(?:etres?)?"], flags=re.I)

        # Keep the useful auction/seller description and equipment list in notes/features.
        # v25.18.3 - build a clean, readable boat equipment list from Grays labels.
        # Grays often flattens the accessories table into one long line, so collecting
        # surrounding text can miss individual items. Detect each known accessory
        # independently and store a concise feature name.
        feature_lines=[]
        accessory_map = (
            (r"\bbimini\b", "Bimini"),
            (r"\bwindscreen\b", "Windscreen"),
            (r"\bbow\s*rail(?:s)?\b", "Bow Rail"),
            (r"\bbait\s*board(?:s)?\b", "Bait Board"),
            (r"\brod\s*holder(?:s)?\b", "Rod Holders"),
            (r"\blive\s*/?\s*kill\s*bait\s*tank\b|\blive.?bait\s*tank\b", "Live / Kill Bait Tank"),
            (r"\bfuel\s*tank\b", "Fuel Tank"),
            (r"\bbatter(?:y|ies)\b", "Battery"),
            (r"\brear\s*(?:end\s*)?seats?\b", "Rear Seats"),
            (r"\banchor\b", "Anchor"),
            (r"\bchain\b", "Chain"),
            (r"\brope\b", "Rope"),
            (r"\bgarmin\b", "Garmin Navigation"),
            (r"\bVHF\s*/?\s*27\s*Meg\b|\bVHF\b|\b27\s*Meg\b", "VHF / 27 Meg Radio"),
            (r"\bwinch\b", "Winch"),
            (r"\bjockey\s*wheel\b", "Jockey Wheel"),
            (r"\b(?:trailer\s*)?lights?\b", "Trailer Lights"),
            (r"\bspare\s*wheel\b", "Spare Wheel"),
            (r"\bsounder\b", "Sounder"),
            (r"\bGPS\b", "GPS"),
        )
        for pattern, label in accessory_map:
            if re.search(pattern, source_text or "", re.I):
                feature_lines.append(label)
        mechanical_lines = []
        if engine_type: mechanical_lines.append("Engine Type: " + engine_type.strip())
        if engine_sn: mechanical_lines.append("Engine Serial: " + engine_sn.strip())
        if engine_turns: mechanical_lines.append("Engine Turns Over: " + engine_turns.strip())
        if beam: mechanical_lines.append("Beam: " + str(beam).strip() + " m")
        if depth: mechanical_lines.append("Depth: " + str(depth).strip() + " m")
        if rego_expiry: mechanical_lines.append("Rego Expiry: " + rego_expiry.strip())
        combined_features = mechanical_lines + feature_lines
        if combined_features: details["boat_features"] = "\n".join(combined_features)[:6000]
        desc=(description or "").strip()
        if desc:
            existing=(details.get("condition_notes") or "").strip()
            if desc.lower() not in existing.lower():
                details["condition_notes"] = ((existing+"\n\n" if existing else "")+"LISTING DESCRIPTION\n"+desc)[:8000]

    # Caravan / camper / trailer specifications (Grays, Slattery and generic listings).
    caravan_probe = " ".join((title or "", description or "", source_text[:50000]))
    if re.search(r"\b(caravan|camper\s*trailer|pop\s*top|motorhome)\b", caravan_probe, re.I):
        details["asset_type"] = "Caravan"
    elif re.search(r"\btrailer\b", caravan_probe, re.I) and not details.get("asset_type"):
        details["asset_type"] = "Trailer"

    def _kg_field(patterns):
        raw = _first_match(source_text, patterns, flags=re.I | re.S)
        v = _number(raw)
        return v if v is not None and 0 < v < 100000 else None

    tare = _kg_field([
        r"(?:tare\s+(?:weight|mass)|tare)\s*[:\-]?\s*([\d,]+(?:\.\d+)?)\s*(?:kg|kgs|kilograms?)?",
    ])
    atm = _kg_field([
        r"(?:aggregate\s+trailer\s+mass|\bATM\b)\s*[:\-]?\s*([\d,]+(?:\.\d+)?)\s*(?:kg|kgs)?",
    ])
    gtm = _kg_field([
        r"(?:gross\s+trailer\s+mass|\bGTM\b|\bGVM\b)\s*[:\-]?\s*([\d,]+(?:\.\d+)?)\s*(?:kg|kgs)?",
    ])
    ball = _kg_field([
        r"(?:tow\s*ball\s+(?:weight|mass)|ball\s+(?:weight|mass)|ball\s+loading)\s*[:\-]?\s*([\d,]+(?:\.\d+)?)\s*(?:kg|kgs)?",
    ])
    if tare is not None: details["tare_weight_kg"] = tare
    if atm is not None: details["atm_kg"] = atm
    if gtm is not None: details["gtm_kg"] = gtm
    if ball is not None: details["ball_weight_kg"] = ball

    berths = _first_match(source_text, [r"(?:berths?|sleeping\s+capacity|sleeps?)\s*[:\-]?\s*(\d{1,2})"], flags=re.I)
    if berths:
        details["berths"] = int(berths)
    axles = _first_match(source_text, [r"(?:axles?|axle\s+configuration)\s*[:\-]?\s*(\d{1,2})"], flags=re.I)
    if axles:
        details["axles"] = int(axles)

    # Dimensions: accept metres or millimetres and normalise to metres.
    def _metres(label):
        m = re.search(rf"(?:{label})\s*[:\-]?\s*([\d,.]+)\s*(mm|cm|m|metres?|meters?)\b", source_text, re.I)
        if not m: return None
        try: v=float(m.group(1).replace(',', ''))
        except ValueError: return None
        unit=m.group(2).lower()
        if unit=='mm': v/=1000
        elif unit=='cm': v/=100
        return round(v,3) if 0 < v < 100 else None
    for key,label in (("length_m",r"overall\s+length|length"),("width_m",r"overall\s+width|width"),("height_m",r"overall\s+height|height")):
        v=_metres(label)
        if v is not None: details[key]=v

    details = _apply_site_specific_details(details, source_text, url=url, title=title, description=description)

    # v25.18.2 - final Grays boat pass.  Run this after the generic/site-specific
    # cleanup so Grays' flattened label/value text cannot overwrite boat fields.
    host = (urllib.parse.urlparse(url or "").netloc or "").lower()
    if "grays.com" in host and details.get("asset_type") == "Boat":
        flat = re.sub(r"\s+", " ", source_text or " " ).strip()

        def grays_boat_value(label, stops):
            stop_alt = "|".join(re.escape(x) for x in stops)
            m = re.search(rf"\b{label}\s*[:\-]?\s*(.+?)(?=\s+(?:{stop_alt})\s*[:\-]?|$)", flat, re.I)
            return re.sub(r"\s+", " ", m.group(1)).strip(" :;,-") if m else ""

        # The Grays page exposes these as labelled fields.
        boat_make = grays_boat_value(r"Make", ["Model", "Asset Sub Category", "Boat Details", "HIN"])
        boat_model = grays_boat_value(r"Model", ["Asset Sub Category", "Boat Details", "HIN", "Rego"])
        if boat_make and len(boat_make) <= 40:
            details["make"] = boat_make
        if boat_model:
            boat_model = re.split(r"\s+(?:Auction\b|\|\s*Grays\b|Grays Australia\b)", boat_model, maxsplit=1, flags=re.I)[0].strip(" -|,")
            if boat_model and len(boat_model) <= 80:
                details["model"] = boat_model

        # Also clean model text derived from the HTML title, e.g.
        # "Seaway 429 Auction (0001-...) | Grays Australia".
        if details.get("model"):
            cleaned_model = re.split(r"\s+(?:Auction\b|\|\s*Grays\b|Grays Australia\b)", str(details["model"]), maxsplit=1, flags=re.I)[0].strip(" -|,")
            if cleaned_model:
                details["model"] = cleaned_model

        hin = grays_boat_value(r"HIN", ["Rego", "State", "Rego Expiry", "Sold Unregistered", "Beam", "Length"])
        rego = grays_boat_value(r"Rego", ["State", "Rego Expiry", "Sold Unregistered", "Beam", "Length"])
        if hin and re.fullmatch(r"[A-Z0-9-]{6,30}", hin, re.I): details["vin"] = hin.upper()
        if rego and len(rego) <= 20: details["registration"] = rego.upper()

        eng_make = grays_boat_value(r"Engine Make", ["Horsepower", "Engine Type", "Fuel Type", "Engine SN", "Engine Hours", "Engine Turns Over"])
        hp = _first_match(flat, [r"Horsepower\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*HP?\b"], flags=re.I)
        eng_type = grays_boat_value(r"Engine Type", ["Fuel Type", "Engine SN", "Engine Hours", "Engine Turns Over", "Boat Accessories"])
        fuel = grays_boat_value(r"Fuel Type", ["Engine SN", "Engine Hours", "Engine Turns Over", "Boat Accessories"])
        eng_sn = grays_boat_value(r"Engine SN", ["Engine Hours", "Engine Turns Over", "Boat Accessories"])
        hours = _first_match(flat, [r"Engine Hours\s*[:\-]?\s*(?:Not\s+sure\s+)?(\d+(?:\.\d+)?)"], flags=re.I)
        if eng_make and len(eng_make) <= 50: details["engine_make"] = eng_make
        if hp: details["horsepower"] = _number(hp)
        if hours: details["engine_hours"] = _number(hours)
        if fuel:
            fuel_low = fuel.lower()
            if "2 stroke" in fuel_low: details["fuel_type"] = "2 Stroke"
            elif "4 stroke" in fuel_low: details["fuel_type"] = "4 Stroke"
            elif "diesel" in fuel_low: details["fuel_type"] = "Diesel"
            elif "petrol" in fuel_low or "gasoline" in fuel_low: details["fuel_type"] = "Petrol"
            else: details["fuel_type"] = "Other"

        beam = _first_match(flat, [r"Beam\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*m"], flags=re.I)
        length = _first_match(flat, [r"Length\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*m"], flags=re.I)
        depth = _first_match(flat, [r"Depth\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*m"], flags=re.I)
        rego_expiry = _first_match(flat, [r"Rego Expiry\s*[:\-]?\s*(\d{1,2}/\d{1,2}/\d{4})"], flags=re.I)
        if beam: details["width_m"] = _number(beam)
        if length: details["length_m"] = _number(length)

        # A Grays boat listing that contains trailer details means a trailer is included.
        if re.search(r"\bTrailer Make\b|\bJockey Wheel\b|\bSpare Wheel\b", flat, re.I):
            details["trailer_included"] = 1

        feature_names = [
            "Bimini", "Windscreen", "Bow Rail", "Bait Boards", "Rod Holders",
            "Live/Kill Bait Tank", "Fuel Tank", "Battery", "Rear End Seats",
            "Anchor", "Chain", "Rope", "Garmin", "VHF/27 Meg", "Winch",
            "Jockey Wheel", "Lights", "Spare Wheel"
        ]
        found_features = [name for name in feature_names if re.search(re.escape(name).replace(r"/", r"[/ ]"), flat, re.I)]
        feature_lines = []
        if eng_type: feature_lines.append(f"Engine Type: {eng_type}")
        if fuel: feature_lines.append(f"Fuel Type: {fuel}")
        if eng_sn: feature_lines.append(f"Engine Serial: {eng_sn}")
        if re.search(r"Engine Turns Over\s*[:\-]?\s*Yes", flat, re.I): feature_lines.append("Engine Turns Over: Yes")
        if beam: feature_lines.append(f"Beam: {beam} m")
        if depth: feature_lines.append(f"Depth: {depth} m")
        if rego_expiry: feature_lines.append(f"Rego Expiry: {rego_expiry}")
        feature_lines.extend(found_features)
        if feature_lines:
            details["boat_features"] = "\n".join(dict.fromkeys(feature_lines))[:6000]

        # Preserve a concise inspection record even when Grays' meta description is sparse.
        note_lines = []
        if description and description.strip(): note_lines.append(description.strip())
        if rego_expiry: note_lines.append(f"Rego Expiry: {rego_expiry}")
        if re.search(r"\bSold Unregistered\b", flat, re.I): note_lines.append("Sold Unregistered")
        if hours and re.search(r"Engine Hours\s*[:\-]?\s*Not\s+sure", flat, re.I): note_lines.append(f"Engine Hours: Not sure {hours}")
        if eng_sn: note_lines.append(f"Engine Serial: {eng_sn}")
        if re.search(r"Trailer Has No VIN", flat, re.I): note_lines.append("Trailer has no VIN - check registration requirements.")
        existing = (details.get("condition_notes") or "").strip()
        combined_notes = [existing] if existing else []
        combined_notes.extend(x for x in note_lines if x and x.lower() not in existing.lower())
        if combined_notes: details["condition_notes"] = "\n\n".join(combined_notes)[:8000]

    # v25.13.1 - Caravan equipment / fit-out notes.
    # Grays and Slattery often place the useful caravan equipment in the lot
    # description rather than in structured fields. Keep that information in
    # both Caravan Features and Condition / Inspection Notes.
    if details.get("asset_type") == "Caravan":
        host = (urllib.parse.urlparse(url or "").netloc or "").lower()
        if "grays.com" in host or "slattery" in host:
            equipment_terms = re.compile(
                r"\b(fridge|refrigerator|solar|solar panel|battery|batteries|inverter|air\s*condition|"
                r"aircon|air con|hot\s*water|toilet|shower|ensuite|microwave|awning|annex|tv|television|"
                r"gas\s*bottle|water\s*tank|fresh\s*water|grey\s*water|bed|bunk|sleeps?|berth|"
                r"generator|reversing\s*camera|reverse\s*camera|suspension|brakes?|stove|cooktop|oven|"
                r"rangehood|washing\s*machine|stereo|radio|antenna|aerial|12v|240v|charger|toolbox|"
                r"bike\s*rack|jerry\s*can|stone\s*guard|sway\s*control|stabiliser|stabilizer)\b", re.I)

            # Prefer the listing description because it avoids auction-site navigation text.
            candidates = []
            for block in (description or "",):
                block = re.sub(r"\s+", " ", block).strip()
                if block and equipment_terms.search(block):
                    candidates.append(block)

            # Also collect useful labelled/equipment lines from the flattened page.
            for m in re.finditer(r"[^\r\n]{0,180}(?:fridge|refrigerator|solar|batter(?:y|ies)|inverter|air\s*condition(?:ing)?|aircon|hot\s*water|ensuite|toilet|shower|microwave|awning|annex|water\s*tank|gas\s*bottle|generator|stove|cooktop|oven|washing\s*machine)[^\r\n]{0,260}", source_text or "", re.I):
                line = re.sub(r"\s+", " ", m.group(0)).strip(" -:;,.|")
                if 5 <= len(line) <= 500:
                    candidates.append(line)

            # De-duplicate while preserving order.
            useful = []
            seen = set()
            for part in candidates:
                key = re.sub(r"[^a-z0-9]+", " ", part.lower()).strip()
                if key and key not in seen:
                    seen.add(key)
                    useful.append(part)

            if useful:
                equipment_text = "\n".join(useful)[:6000]
                details["caravan_features"] = equipment_text
                existing = (details.get("condition_notes") or "").strip()
                heading = "CARAVAN EQUIPMENT / FEATURES"
                if heading.lower() not in existing.lower():
                    combined = (existing + "\n\n" if existing else "") + heading + "\n" + equipment_text
                    details["condition_notes"] = combined[:8000]

    return details


def _apply_site_specific_details(details, source_text, url="", title="", description=""):
    """Clean up fields for major Australian listing/auction sites.

    Generic parsing remains the fallback, but each supported site gets stricter
    rules so navigation text is not mistaken for vehicle data.
    """
    host = (urllib.parse.urlparse(url or "").netloc or "").lower()
    text = source_text or ""
    td = " ".join(x for x in (title, description) if x)
    slug = urllib.parse.unquote((urllib.parse.urlparse(url or "").path or "").replace("-", " "))
    vehicle_probe = " ".join(x for x in (td, slug) if x)

    def set_if(key, value):
        if value not in (None, ""):
            details[key] = value


    # v25.11.3: reject JavaScript field-name tokens that can look like real values.
    bad_tokens = {"expiry", "expirydate", "registrationexpiry", "registrationdate", "odometer", "capacity", "fueltype", "bodytype", "transmission", "drivetype", "colour", "color", "vin"}
    def plausible_token(value):
        return bool(value) and re.sub(r"[^a-z0-9]", "", str(value).lower()) not in bad_tokens

    def clean_site_note(*bad_phrases):
        note = (details.get("condition_notes") or "").strip()
        low = note.lower()
        if note and any(p.lower() in low for p in bad_phrases):
            details.pop("condition_notes", None)

    # Common auction finish wording, including Manheim-style natural language.
    if any(x in host for x in ("manheim", "pickles", "slattery")):
        finish_text = _first_match(text, [
            r"(?:auction\s+ends?|auction\s+finishes?|closing\s+time|closes?)\s*[:\-]?\s*(?:on\s+)?((?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)?[,]?\s*\d{1,2}\s+[A-Za-z]+\s+20\d{2}\s+\d{1,2}(?::\d{2})?\s*(?:AM|PM)?)",
            r"(?:for\s+auction\s+in.*?\s+on\s+)((?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)[,]?\s*\d{1,2}\s+[A-Za-z]+\s+20\d{2}\s+\d{1,2}(?::\d{2})?\s*(?:AM|PM)?)",
        ], flags=re.I | re.S)
        if finish_text:
            dt = _auction_datetime_value(finish_text)
            if dt:
                details["auction_finish"] = dt

    # v25.11.4: build exact label/value pairs before site-specific mapping.
    # This prevents labels such as NUMBER / EXPIRYDATE from being treated as values.
    def labelled_value(label, stop_labels=()):
        stops = list(stop_labels) + [
            "Make", "Model", "Colour", "Color", "VIN", "Series", "Transmission",
            "Drive Type", "Body Type", "Registration Number", "Registration Expiry",
            "Engine Number", "Engine Type", "Odometer", "Fuel Type", "Capacity",
            "No of Seats", "No. of Seats", "Owners Manual", "Service History",
            "Compliance Date", "Build Date", "Trim", "Cylinders", "Registration",
            "GVM", "GCM", "KW RPM", "Track", "Width", "Height", "Log Book",
            "Wheelbase", "Kerb Weight", "Total Quantity", "Sold Registered"
        ]
        stop_alt = "|".join(re.escape(x) for x in sorted(set(stops), key=len, reverse=True))
        pat = rf"(?:^|\s){re.escape(label)}\s*[:\-]?\s*(.+?)(?=\s+(?:{stop_alt})\b|$)"
        m = re.search(pat, text, flags=re.I | re.S)
        return re.sub(r"\s+", " ", m.group(1)).strip(" :;-.,") if m else ""

    def jsonish_value(*keys):
        # Read simple values embedded in flattened JavaScript/JSON state.
        for key in keys:
            patterns = [
                rf"(?i)(?:[\"'\\]?){re.escape(key)}(?:[\"'\\]?)\s*[:=]\s*(?:\\?[\"']?)\s*([^,}}\]\r\n]{{1,120}})",
                rf"(?i)\b{re.escape(key)}\b\s*[:\\]+\s*(?:\\?[\"']?)\s*([^,}}\]\r\n]{{1,120}})",
            ]
            for pat2 in patterns:
                m2 = re.search(pat2, text or "")
                if m2:
                    v = m2.group(1).strip().strip("\\\"' :;")
                    if v:
                        return re.sub(r"\s+", " ", v)
        return ""

    def jsonish_number(*keys):
        v = jsonish_value(*keys)
        m = re.search(r"-?\d+(?:\.\d+)?", v or "")
        return m.group(0) if m else ""

    if "manheim" in host:
        details["auction_name"] = "Manheim"
        # Manheim's numeric asset ID is stable and preferable to stray state text.
        m = re.search(r"/passenger-vehicles/(\d+)", urllib.parse.urlparse(url).path, flags=re.I)
        if m:
            details["lot_number"] = m.group(1)
        elif (details.get("lot_number") or "").upper() in {"VIC", "NSW", "QLD", "SA", "WA", "TAS", "NT", "ACT"}:
            details.pop("lot_number", None)

        loc = details.get("auction_location") or ""
        if any(word in loc.lower() for word in ("site links", "price guide", "about us", "sign in", "locations")):
            details.pop("auction_location", None)
        manheim_loc = _first_match(text, [
            r"(?:auction\s+location|location|site)\s*[:\-]?\s*([A-Za-z][A-Za-z .'-]+,\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b",
        ])
        if manheim_loc:
            details["auction_location"] = re.sub(r"\s+", " ", manheim_loc).strip()

        # Ranger/utility pages commonly encode diesel as 3.2D / 2.0D in title or URL.
        if re.search(r"\b(?:diesel|\d(?:\.\d)?\s*d\b|\d(?:\.\d)?dt\b)", vehicle_probe, flags=re.I):
            details["fuel_type"] = "Diesel"
        elif re.search(r"\bpetrol\b", vehicle_probe, flags=re.I):
            details["fuel_type"] = "Petrol"
        elif details.get("fuel_type") in {"Hybrid", "Plug-in Hybrid", "Electric"} and not re.search(r"\b(?:hybrid|phev|electric|ev)\b", vehicle_probe, flags=re.I):
            details.pop("fuel_type", None)

        size = _first_match(vehicle_probe, [r"\b(\d(?:\.\d)?)\s*d(?:t)?\b"])
        if size:
            details["engine_size"] = f"{size}L"
        if re.search(r"\bdual\s+cab\s+utility\b", vehicle_probe, flags=re.I):
            details["body_type"] = "Dual Cab Utility"
        elif re.search(r"\butility\b", vehicle_probe, flags=re.I):
            details["body_type"] = "Utility"
        clean_site_note("for auction in national online", "used 2016 ford ranger")

    elif "pickles" in host:
        details["auction_name"] = "Pickles"

        # Pickles vehicle-detail tables use clear label/value pairs. Pull those
        # values directly instead of relying on broad page-wide guesses.
        pickles_odo = _first_match(text, [
            r"Odometer\s*\(Showing\s+on\)\s*[:\-]?\s*([\d,]+)\s*(?:km|kms)\b",
            r"Odometer(?:\s*\([^)]*\))?\s*[:\-]?\s*([\d,]{3,})\s*(?:km|kms)\b",
            r"\b([\d,]{4,})\s*(?:km|kms)\b",
        ])
        v = _number(pickles_odo, integer=True)
        if v is not None:
            details["odometer_km"] = v

        pickles_vin = _first_match(text, [r"\bVIN\s*[:\-]?\s*([A-HJ-NPR-Z0-9]{17})\b"])
        if pickles_vin:
            details["vin"] = pickles_vin.upper()

        pickles_colour = _first_match(text, [
            r"(?:Keys|Spare\s+Keys)\s+(?:[^ ]+\s+){0,4}(White\s*-\s*Artic\s+White|[A-Za-z]+\s*-\s*[A-Za-z ]{2,30})(?=\s+(?:Compliance\s+Date|Build\s+Date))",
            r"(?:Colour|Color)\s*[:\-]?\s*([A-Za-z][A-Za-z /-]{1,40}?)(?=\s+(?:Compliance\s+Date|Build\s+Date|Odometer))",
        ])
        if pickles_colour:
            details["colour"] = re.sub(r"\s+", " ", pickles_colour).strip()

        pickles_trans = _first_match(text, [
            r"Transmission\s*[:\-]?\s*((?:\d+\s*Spd\s*)?(?:Sports\s+Automatic|Automatic|Manual|CVT|DCT))\b",
        ])
        if pickles_trans:
            details["transmission"] = "Sports Automatic" if re.search(r"sports\s+automatic", pickles_trans, re.I) else re.sub(r"^\d+\s*Spd\s*", "", pickles_trans, flags=re.I).strip().title()

        pickles_drive = _first_match(text, [r"Drive\s+Type\s*[:\-]?\s*([^|]{2,40}?)(?=\s+Engine\s+Capacity|\s+Fuel\b|$)"])
        if pickles_drive:
            pd = pickles_drive.lower()
            if "4x4" in pd or "4wd" in pd or "four wheel" in pd:
                details["drive_type"] = "4WD"
            elif "awd" in pd or "all wheel" in pd:
                details["drive_type"] = "AWD"
            elif "2wd" in pd:
                details["drive_type"] = "2WD"

        pickles_engine = _first_match(text, [r"Engine\s+Capacity\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*(?:Ltr|Litre|Liter|L)?\b"])
        if pickles_engine:
            details["engine_size"] = f"{pickles_engine}L"
            try:
                details["engine_cc"] = int(round(float(pickles_engine) * 1000))
            except ValueError:
                pass

        pickles_cyl = _first_match(text, [
            r"Cylinders\s*[:\-]?\s*(\d{1,2})\b",
            r"\b(\d{1,2})\s+cyl\b",
        ])
        if pickles_cyl:
            details["engine_cylinders"] = pickles_cyl
        if not details.get("engine_size"):
            search_litre = _first_match(text, [r"\b(\d+(?:\.\d+)?)\s+L\s+(?:Diesel|Petrol)\b"])
            if search_litre:
                details["engine_size"] = f"{search_litre}L"
                try:
                    details["engine_cc"] = int(round(float(search_litre) * 1000))
                except ValueError:
                    pass

        if re.search(r"Fuel\s*[:\-]?\s*(?:Direct\s+Injection\s+)?Diesel\b", text, re.I) or re.search(r"\b\d+(?:\.\d+)?\s+L\s+Diesel\b", text, re.I):
            details["fuel_type"] = "Diesel"
        elif re.search(r"Fuel\s*[:\-]?\s*(?:Unleaded\s+)?Petrol\b", text, re.I) or re.search(r"\b\d+(?:\.\d+)?\s+L\s+Petrol\b", text, re.I):
            details["fuel_type"] = "Petrol"

        pickles_trim = _first_match(text, [r"Trim\s*[:\-]?\s*([^|]{2,60}?)(?=\s+Registration\b|\s+No\.\s*of\s+Seats|$)"])
        if pickles_trim:
            details["interior"] = re.sub(r"\s+", " ", pickles_trim).strip(" ,.-")

        pickles_seats = _first_match(text, [
            r"No\.\s*of\s+Seats\s*[:\-]?\s*(\d{1,2})\b",
            r"\b(\d{1,2})\s+seats?\b",
        ])
        if pickles_seats:
            details["seat_count"] = int(pickles_seats)

        if re.search(r"Registration\s*[:\-]?\s*No\s+Registration\b", text, re.I):
            details["registration"] = ""
            details["registration_status"] = "Unregistered"
        else:
            preg_candidates = re.findall(r"Registration\s*[:\-]?\s*([A-Z0-9-]{3,12})\b", text, flags=re.I)
            preg = next((x for x in preg_candidates if plausible_token(x) and x.lower() != "no"), "")
            if preg:
                details["registration"] = preg.upper()
                details.setdefault("registration_status", "Registered")
            elif not plausible_token(details.get("registration")):
                details["registration"] = ""

        # Explicit LOT text is the correct auction lot. STOCK is a separate Pickles ID.
        explicit_lot = _first_match(text, [r"\bLOT\s+(\d{1,8})\b"])
        if explicit_lot:
            details["lot_number"] = explicit_lot
        loc = _first_match(text, [
            r"\blocated\s+at\s+([A-Za-z][A-Za-z .'-]+,\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b",
            r"(?:location|branch)\s*[:\-]?\s*([A-Za-z][A-Za-z .'-]+,\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b",
            r"\b([A-Za-z][A-Za-z .'-]+,\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b(?=\s+(?:Pickles|Add a note|National|Sale Info))",
        ])
        if loc:
            details["auction_location"] = re.sub(r"\s+", " ", loc).strip()
        elif details.get("auction_location") and any(x in details["auction_location"].lower() for x in ("contact us", "media complaints", "technical")):
            details.pop("auction_location", None)

        # Pickles description shorthand: Spts Auto, 4x4, 3.2DT, Pick-up/Super Cab.
        if re.search(r"\bspts?\s+auto\b", vehicle_probe, flags=re.I):
            details["transmission"] = "Sports Automatic"
        if re.search(r"\b4x4\b|\b4wd\b", vehicle_probe, flags=re.I):
            details["drive_type"] = "4WD"
        eng = _first_match(vehicle_probe, [r"\b(\d(?:\.\d)?)\s*DT\b", r"\b(\d(?:\.\d)?)\s*D\b"])
        if eng:
            details["engine_size"] = f"{eng}L"
            details["fuel_type"] = "Diesel"
        elif re.search(r"\bdiesel\b", vehicle_probe, flags=re.I):
            details["fuel_type"] = "Diesel"
        elif re.search(r"\bpetrol\b", vehicle_probe, flags=re.I):
            details["fuel_type"] = "Petrol"
        elif not re.search(r"\b(?:hybrid|electric|phev|ev)\b", vehicle_probe, flags=re.I):
            details.pop("fuel_type", None)
        if re.search(r"\bpick[- ]?up\b", vehicle_probe, flags=re.I):
            details["body_type"] = "Pick-up"
        elif re.search(r"\b(?:ute|utility)\b", vehicle_probe, flags=re.I):
            details["body_type"] = "Utility"
        # Exact Pickles table overrides.
        pv = labelled_value("Odometer (Showing on)") or labelled_value("Odometer")
        m = re.search(r"([\d,]+)\s*(?:km|kms)\b", pv, re.I)
        if m: details["odometer_km"] = int(m.group(1).replace(",", ""))
        pv = labelled_value("Engine Capacity")
        m = re.search(r"(\d+(?:\.\d+)?)", pv)
        if m: details["engine_size"] = f"{m.group(1)}L"
        pv = labelled_value("Fuel")
        if re.search(r"diesel", pv, re.I): details["fuel_type"] = "Diesel"
        elif re.search(r"petrol", pv, re.I): details["fuel_type"] = "Petrol"
        pv = labelled_value("Registration")
        if re.search(r"no\s+registration", pv, re.I):
            details["registration"] = ""; details["registration_status"] = "Unregistered"
        elif pv and plausible_token(pv): details["registration"] = pv.upper()
        if (details.get("registration") or "").lower() in {"number", "expiry", "expirydate"}: details["registration"] = ""

        # v25.11.5 - exact Pickles client-state fallbacks.
        raw_reg = jsonish_value("registrationNumber", "regoNumber", "registrationNo")
        if raw_reg and raw_reg.lower() not in {"null", "none", "undefined", "false"}:
            m = re.search(r"\b([A-Z0-9-]{3,10})\b", raw_reg, re.I)
            if m and plausible_token(m.group(1)):
                details["registration"] = m.group(1).upper()
                details["registration_status"] = "Registered"
        reg_now = str(details.get("registration") or "")
        if any(tok in reg_now.lower() for tok in ("expirydate", "registrationdate", "registr", "odometer", "capacity")):
            details["registration"] = ""
        if re.search(r"\bNo\s+Registration\b", text, re.I) or re.search(r"registration(?:Number|No)?\s*[:=]\s*(?:null|none|undefined|false|\"\")", text, re.I):
            details["registration"] = ""
            details["registration_status"] = "Unregistered"
        elif not details.get("registration") and details.get("registration_status") == "Registered":
            details.pop("registration_status", None)

        jodo = jsonish_number("odometer", "odometerKm", "odometerReading", "kilometres", "kilometers")
        if jodo:
            try:
                km = int(float(jodo))
                if 100 <= km < 10_000_000: details["odometer_km"] = km
            except ValueError:
                pass

        jeng = jsonish_number("engineCapacity", "engineSize", "capacity")
        if jeng:
            try:
                ev = float(jeng)
                if ev >= 100:
                    details["engine_cc"] = int(round(ev))
                    details["engine_size"] = f"{ev/1000.0:.1f}L"
                elif 0.5 <= ev <= 10:
                    details["engine_size"] = f"{ev:g}L"
                    details["engine_cc"] = int(round(ev * 1000))
            except ValueError:
                pass

        jf = jsonish_value("fuelType", "fuel")
        if re.search(r"diesel", jf, re.I): details["fuel_type"] = "Diesel"
        elif re.search(r"petrol|unleaded", jf, re.I): details["fuel_type"] = "Petrol"

        jb = jsonish_value("bodyType", "bodyStyle")
        if jb and plausible_token(jb):
            details["body_type"] = re.sub(r"[^A-Za-z0-9 /-].*$", "", jb).strip().title()
        elif re.search(r"\bPick[- ]?up\b", text, re.I):
            details["body_type"] = "Pick-up"

        clean_site_note("located at sunshine", "buy 2022 ford ranger")

    elif "slattery" in host:
        details["auction_name"] = "Slattery Auctions"

        # Slattery exposes an Item Details table. These labelled values are more
        # reliable than generic matches elsewhere on the auction page.
        sy = _first_match(text, [r"Year\s+Of\s+Manufacture\s*[:\-]?\s*((?:19|20)\d{2})\b"])
        if sy:
            details["year"] = int(sy)

        smake = _first_match(text, [r"\bMake\s*[:\-]?\s*([A-Za-z][A-Za-z0-9 .'-]{1,30}?)(?=\s+Model\b|\s+Colour\b)"])
        if smake:
            details["make"] = re.sub(r"\s+", " ", smake).strip()
        smodel = _first_match(text, [r"\bModel\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9 .'-]{1,40}?)(?=\s+Colour\b|\s+VIN\b|\s+Series\b)"])
        if smodel:
            details["model"] = re.sub(r"\s+", " ", smodel).strip()

        scolour = _first_match(text, [r"\bColour\s*[:\-]?\s*([A-Za-z][A-Za-z /-]{1,30}?)(?=\s+VIN\b|\s+Series\b)"])
        if scolour:
            details["colour"] = re.sub(r"\s+", " ", scolour).strip()
        svin = _first_match(text, [r"\bVIN\s*[:\-]?\s*([A-HJ-NPR-Z0-9]{17})\b"])
        if svin:
            details["vin"] = svin.upper()
        sreg_candidates = re.findall(r"Registration\s+Number\s*[:\-]?\s*([A-Z0-9-]{3,12})\b", text, flags=re.I)
        sreg = next((x for x in sreg_candidates if plausible_token(x) and not x.isdigit()), "")
        if sreg:
            details["registration"] = sreg.upper()
        elif not plausible_token(details.get("registration")):
            details["registration"] = ""

        strans = _first_match(text, [r"Transmission\s*[:\-]?\s*(Sports\s+Automatic|Automatic|Manual|CVT|DCT)\b"])
        if strans:
            details["transmission"] = strans.title() if strans.lower() != "sports automatic" else "Sports Automatic"

        sdrive = _first_match(text, [r"Drive\s+Type\s*[:\-]?\s*(Four\s+Wheel\s+Drive|All\s+Wheel\s+Drive|Front\s+Wheel\s+Drive|Rear\s+Wheel\s+Drive|4WD|AWD|2WD)\b"])
        if sdrive:
            dl = sdrive.lower()
            details["drive_type"] = "4WD" if ("four" in dl or "4wd" in dl) else ("AWD" if ("all" in dl or "awd" in dl) else ("2WD" if "2wd" in dl else ("FWD" if "front" in dl else "RWD")))

        sbody = _first_match(text, [r"Body\s+Type\s*[:\-]?\s*(SUV|Sedan|Wagon|Hatchback|Utility|Ute|Coupe|Van|Bus|Truck|Convertible|Cab Chassis|Dual Cab(?: Utility)?)\b"])
        if sbody:
            details["body_type"] = re.sub(r"\s+", " ", sbody).strip().upper() if len(sbody.strip()) <= 4 else re.sub(r"\s+", " ", sbody).strip().title()

        sodo = _first_match(text, [r"Odometer(?:\s*\([^)]*\))?\s*[:\-]?\s*([\d,]{3,})\s*(?:KMs?|Kilometres?|KM)\b"])
        ov = _number(sodo, integer=True)
        if ov is not None:
            details["odometer_km"] = ov

        sfuel = _first_match(text, [r"Fuel\s+Type\s*[:\-]?\s*(Diesel|Petrol|Electric|Hybrid|LPG)\b"])
        if sfuel:
            details["fuel_type"] = sfuel.title()

        scc = _first_match(text, [r"\bCapacity\s*[:\-]?\s*([\d,]{3,5})(?:\s*(?:cc|CC))?\b"])
        ccv = _number(scc, integer=True)
        if ccv:
            details["engine_cc"] = ccv
            details["engine_size"] = f"{ccv/1000.0:.1f}L"

        sseats = _first_match(text, [r"No\s+of\s+Seats\s*[:\-]?\s*(\d{1,2})\b"])
        if sseats:
            details["seat_count"] = int(sseats)

        # Slattery represents the sale-registration flag as a labelled boolean.
        # Preserve the useful business meaning in BAM instead of storing True/False.
        if re.search(r"Sold\s+Registered,?\s+Sold\s+on\s+Consignment\s*[:\-]?\s*True\b", text, re.I):
            details["registration_status"] = "Sold Registered, Sold on Consignment"
        elif re.search(r"Sold\s+Registered\s*[:\-]?\s*True\b", text, re.I):
            details["registration_status"] = "Sold Registered"

        loc = details.get("auction_location") or ""
        if any(x in loc.lower() for x in ("category", "contact", "about", "services")):
            details.pop("auction_location", None)
        better_loc = _first_match(text, [
            r"(?:location|yard|branch)\s*[:\-]?\s*([A-Za-z][A-Za-z .'-]+,\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b",
        ])
        if better_loc:
            details["auction_location"] = re.sub(r"\s+", " ", better_loc).strip()
        # Exact Slattery Item Details overrides.
        exact_map = {
            "Registration Number": "registration", "Body Type": "body_type",
            "Fuel Type": "fuel_type", "Transmission": "transmission"
        }
        for lab, key in exact_map.items():
            val = labelled_value(lab)
            if val and plausible_token(val): details[key] = val
        val = labelled_value("Odometer")
        m = re.search(r"([\d,]+)\s*(?:km|kms)", val, re.I)
        if m: details["odometer_km"] = int(m.group(1).replace(",", ""))
        val = labelled_value("Capacity")
        m = re.search(r"([\d,]{3,5})", val)
        if m:
            cc=int(m.group(1).replace(",", "")); details["engine_cc"]=cc; details["engine_size"]=f"{cc/1000:.3f}L".rstrip("0").rstrip(".")+"L" if False else f"{cc/1000:.3f}L".rstrip("0").rstrip(".")
        val = labelled_value("No of Seats")
        m = re.search(r"\d{1,2}", val)
        if m: details["seat_count"] = int(m.group())
        sold = labelled_value("Sold Registered, Sold on Consignment")
        if sold.lower().startswith("true"): details["registration_status"] = "Sold Registered, Sold on Consignment"

        # v25.11.5 - exact Slattery client-state fallbacks.
        jy = jsonish_number("yearOfManufacture", "manufactureYear")
        if jy and re.fullmatch(r"(?:19|20)\d{2}", jy): details["year"] = int(jy)

        jr = jsonish_value("registrationNumber", "regoNumber")
        if jr:
            m = re.search(r"\b([A-Z0-9-]{3,10})\b", jr, re.I)
            if m and plausible_token(m.group(1)): details["registration"] = m.group(1).upper()

        jo = jsonish_number("odometer", "odometerKm", "kilometres", "kilometers")
        if jo:
            try:
                km = int(float(jo))
                if 100 <= km < 10_000_000: details["odometer_km"] = km
            except ValueError:
                pass

        jfuel = jsonish_value("fuelType", "fuel")
        if re.search(r"diesel", jfuel, re.I): details["fuel_type"] = "Diesel"
        elif re.search(r"petrol", jfuel, re.I): details["fuel_type"] = "Petrol"

        jbody = jsonish_value("bodyType", "bodyStyle")
        if jbody:
            b = re.sub(r"[^A-Za-z0-9 /-].*$", "", jbody).strip()
            if b: details["body_type"] = b.upper() if len(b) <= 4 else b.title()

        jseats = jsonish_number("noOfSeats", "numberOfSeats", "seatCount", "seats")
        if jseats:
            try:
                n = int(float(jseats))
                if 1 <= n <= 99: details["seat_count"] = n
            except ValueError:
                pass

        jtrans = jsonish_value("transmission")
        if re.search(r"sports\s+automatic", jtrans, re.I) or re.search(r"sports\s+automatic", text, re.I): details["transmission"] = "Sports Automatic"
        elif re.search(r"automatic", jtrans, re.I): details["transmission"] = "Automatic"
        elif re.search(r"manual", jtrans, re.I): details["transmission"] = "Manual"

        jcap = jsonish_number("capacity", "engineCapacity")
        if jcap:
            try:
                cc = int(round(float(jcap)))
                if 500 <= cc <= 10000:
                    details["engine_cc"] = cc
                    details["engine_size"] = f"{cc/1000.0:.3f}".rstrip("0").rstrip(".") + "L"
            except ValueError:
                pass

        sold_reg = jsonish_value("soldRegistered")
        sold_cons = jsonish_value("soldOnConsignment", "onConsignment")
        if re.search(r"true|1", sold_reg, re.I) and re.search(r"true|1", sold_cons, re.I):
            details["registration_status"] = "Sold Registered, Sold on Consignment"
        elif re.search(r"true|1", sold_reg, re.I):
            details["registration_status"] = "Sold Registered"

        if (details.get("lot_number") or "").lower() in {"number", "expiry", "expirydate"}: details.pop("lot_number", None)
        clean_site_note("professional auctioneers and valuers", "wide variety of general & specialised auction sales")
        if re.search(r"\bdiesel\b", vehicle_probe, flags=re.I):
            details["fuel_type"] = "Diesel"
        elif re.search(r"\bpetrol\b", vehicle_probe, flags=re.I):
            details["fuel_type"] = "Petrol"
        elif details.get("fuel_type") in {"Hybrid", "Plug-in Hybrid", "Electric"} and not re.search(r"\b(?:hybrid|phev|electric|ev)\b", vehicle_probe, flags=re.I):
            details.pop("fuel_type", None)

    elif "gumtree" in host:
        details["listing_source"] = "Gumtree"
        if details.get("auction_location") and not details.get("seller_location"):
            details["seller_location"] = details.pop("auction_location")

        # Gumtree Australia publishes a very regular Listing Info table. Prefer
        # those exact labels over generic page-wide matches so navigation/spec
        # text cannot contaminate vehicle fields.
        gum_loc = _first_match(text, [
            r"(?:^|\s)Location\s+([A-Za-z][A-Za-z .'-]+,\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b",
            r"(?:location|located\s+in)\s*[:\-]?\s*([A-Za-z][A-Za-z .'-]+,\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b",
        ], flags=re.I)
        if gum_loc:
            details["seller_location"] = re.sub(r"\s+", " ", gum_loc).strip()

        seller = _first_match(text, [
            r"(?:^|\s)Listed\s+By\s+(.+?)(?=\s+(?:Views|Last\s+Edited|Date\s+Listed|Listing\s+ID|Private\s+seller|Dealer\s+used)\b|$)",
        ], flags=re.I | re.S)
        if seller and len(seller.strip()) <= 80:
            details["seller_name"] = re.sub(r"\s+", " ", seller).strip(" ,-:")

        exact_patterns = {
            "variant": [r"(?:^|\s)Variant\s+(.+?)(?=\s+Body\s+Type\b)"],
            "body_type": [r"(?:^|\s)Body\s+Type\s+(.+?)(?=\s+Year\b)"],
            "odometer_km": [r"(?:^|\s)Odometer\s+([\d,]+)\s*km\b"],
            "transmission": [r"(?:^|\s)Transmission\s+(.+?)(?=\s+Drive\s+Train\b)"],
            "drive_type": [r"(?:^|\s)Drive\s+Train\s+(.+?)(?=\s+Fuel\s+Type\b)"],
            "fuel_type": [r"(?:^|\s)Fuel\s+Type\s+(.+?)(?=\s+(?:Engine\s+Capacity|Cylinder\s+Configuration|Colour|Color|Air\s+conditioning|Is\s+your\s+car\s+registered)\b)"],
            "engine_size": [r"(?:^|\s)Engine\s+Capacity\s+([\d.]+\s*[Ll]|[\d,]{3,5})\b"],
            "engine_cylinders": [r"(?:^|\s)Cylinder\s+Configuration\s+([0-9]{1,2})(?:\s*cyl)?\b"],
            "colour": [r"(?:^|\s)Colour\s+(.+?)(?=\s+(?:Air\s+conditioning|Is\s+your\s+car\s+registered|Registration\s+number|VIN|Stock\s+Number|Location)\b)"],
            "registration": [r"(?:^|\s)Registration\s+number\s+([A-Z0-9-]{2,10})\b"],
            "vin": [r"(?:^|\s)VIN\s+([A-HJ-NPR-Z0-9]{17})\b"],
        }
        for key, patterns in exact_patterns.items():
            val = _first_match(text, patterns, flags=re.I | re.S)
            if not val:
                continue
            val = re.sub(r"\s+", " ", str(val)).strip(" ,-:")
            if key == "odometer_km":
                num = _number(val, integer=True)
                if num is not None: details[key] = num
            elif key == "engine_size":
                raw = val.replace(",", "").strip()
                if re.fullmatch(r"\d{3,5}", raw):
                    cc = int(raw)
                    details["engine_cc"] = cc
                    details[key] = f"{cc/1000.0:.1f}L"
                else:
                    details[key] = re.sub(r"\s+", "", val).upper().replace("L", "L")
            elif key == "transmission":
                if re.search(r"sports?\s+automatic|spts?\s+auto", val, re.I): details[key] = "Sports Automatic"
                elif re.search(r"automatic", val, re.I): details[key] = "Automatic"
                elif re.search(r"manual", val, re.I): details[key] = "Manual"
            elif key == "drive_type":
                if re.search(r"four\s+wheel|4wd|4x4", val, re.I): details[key] = "4WD"
                elif re.search(r"all[ -]?wheel|awd", val, re.I): details[key] = "AWD"
                elif re.search(r"front[ -]?wheel|fwd", val, re.I): details[key] = "FWD"
                elif re.search(r"rear[ -]?wheel|rwd", val, re.I): details[key] = "RWD"
                else: details[key] = val
            elif key == "fuel_type":
                if re.search(r"diesel", val, re.I): details[key] = "Diesel"
                elif re.search(r"petrol|unleaded", val, re.I): details[key] = "Petrol"
                elif re.search(r"hybrid", val, re.I): details[key] = "Hybrid"
                elif re.search(r"electric", val, re.I): details[key] = "Electric"
                elif re.search(r"lpg", val, re.I): details[key] = "LPG"
            elif key == "colour":
                details[key] = val.title()
            elif key == "vin":
                details[key] = val.upper()
            elif key == "registration":
                details[key] = val.upper()
            else:
                details[key] = val

        reg_yes = _first_match(text, [r"Is\s+your\s+car\s+registered\?\s+(Yes|No)\b"], flags=re.I)
        if reg_yes:
            details["registration_status"] = "Registered" if reg_yes.lower() == "yes" else "Unregistered"
            if reg_yes.lower() == "no":
                details.pop("registration", None)

        listing_id = _first_match(text, [r"Listing\s+ID\s+(\d{6,15})\b"], flags=re.I)
        if listing_id:
            details["lot_number"] = listing_id

        # v25.12.5 - Gumtree copied-text fallbacks. Gumtree's copyable ad text
        # is not always laid out like the live Listing Info table, so scan the
        # pasted advert itself for common dealer/private-sale wording. Keep
        # these rules Gumtree-only so the working auction/Facebook importers
        # are not affected.
        if not details.get("seller_location"):
            gum_loc2 = _first_match(text, [
                r"(?:located\s+(?:at|in)|location|suburb)\s*[:\-]?\s*([A-Za-z][A-Za-z .'-]+,\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b",
                r"\b([A-Za-z][A-Za-z .'-]{2,40},\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b",
            ], flags=re.I)
            if gum_loc2:
                details["seller_location"] = re.sub(r"\s+", " ", gum_loc2).strip()

        if not details.get("seller_phone"):
            gum_phone = _first_match(text, [
                r"(?:phone|mobile|call|contact)\s*[:\-]?\s*((?:\+?61\s*4|04)\d(?:[ \-]?\d){7,8})",
                r"\b((?:\+?61\s*4|04)\d(?:[ \-]?\d){7,8})\b",
            ], flags=re.I)
            if gum_phone:
                details["seller_phone"] = re.sub(r"[^0-9+]", "", gum_phone)

        # Exact or near-exact specification labels commonly present when a
        # Gumtree dealer listing is copied from the browser.
        if not details.get("body_type"):
            gum_body = _first_match(text, [
                r"(?:body\s*(?:type|style)|vehicle\s*type)\s*[:\-]?\s*(SUV|Sedan|Wagon|Hatchback|Hatch|Ute|Utility|Dual Cab|Single Cab|Extra Cab|Cab Chassis|Van|Coupe|Convertible|People Mover|Pickup|Pick-up)",
            ], flags=re.I)
            if gum_body:
                details["body_type"] = gum_body.title()

        if not details.get("seat_count"):
            gum_seats = _first_match(text, [
                r"(?:no\.?\s*of\s*seats|number\s*of\s*seats|seats)\s*[:\-]?\s*(\d{1,2})\b",
                r"\b(\d{1,2})\s*seater\b",
            ], flags=re.I)
            if gum_seats:
                try:
                    n = int(gum_seats)
                    if 1 <= n <= 20:
                        details["seat_count"] = n
                except ValueError:
                    pass

        if not details.get("engine_cylinders"):
            gum_cyl = _first_match(text, [
                r"(?:cylinders?|cylinder\s*configuration)\s*[:\-]?\s*(\d{1,2})\b",
                r"\b(\d{1,2})\s*cyl(?:inder)?s?\b",
            ], flags=re.I)
            if gum_cyl:
                details["engine_cylinders"] = gum_cyl

        if not details.get("fuel_type"):
            if re.search(r"\bdiesel\b", text, flags=re.I):
                details["fuel_type"] = "Diesel"
            elif re.search(r"\b(?:petrol|unleaded)\b", text, flags=re.I):
                details["fuel_type"] = "Petrol"
            elif re.search(r"\bhybrid\b", text, flags=re.I):
                details["fuel_type"] = "Hybrid"
            elif re.search(r"\belectric\b", text, flags=re.I):
                details["fuel_type"] = "Electric"

        if not details.get("registration"):
            gum_reg = _first_match(text, [
                r"(?:registration\s*(?:number|no\.?)?|rego)\s*[:#\-]?\s*([A-Z0-9-]{2,10})\b",
            ], flags=re.I)
            if gum_reg and gum_reg.lower() not in {"status", "number", "expiry", "expires", "date"}:
                details["registration"] = gum_reg.upper()

        if not details.get("registration_status"):
            if re.search(r"\b(?:unregistered|no\s+registration|not\s+registered)\b", text, flags=re.I):
                details["registration_status"] = "Unregistered"
                details.pop("registration", None)
            elif re.search(r"\bregistered\b", text, flags=re.I):
                details["registration_status"] = "Registered"

        # A copied Gumtree description may state interior trim without a formal
        # label. Only map conservative material names supported by BAM.
        if not details.get("interior"):
            gum_int = _first_match(text, [
                r"(?:interior|trim)\s*[:\-]?\s*(black\s*/\s*grey\s+cloth|black\s+cloth|grey\s+cloth|gray\s+cloth|cloth|leather)",
            ], flags=re.I)
            if gum_int:
                details["interior"] = "Leather" if "leather" in gum_int.lower() else "Cloth"

        # Keep the whole copied Gumtree advert in Condition / Inspection Notes,
        # matching the Facebook workflow, while individual fields are extracted
        # above for searching and valuation.
        if text and not (details.get("condition_notes") or "").strip():
            gum_note = _strip_html(text).strip()
            if gum_note and not gum_note.lower().startswith(("http://", "https://")):
                details["condition_notes"] = gum_note[:4000]

        # Gumtree listings are fixed-price marketplace ads, never auctions.
        if details.get("current_bid") and not details.get("asking_price"):
            details["asking_price"] = details.pop("current_bid")
        details.pop("auction_location", None)
        details.pop("auction_name", None)

    elif "facebook.com" in host or "fb.com" in host:
        details["listing_source"] = "Facebook Marketplace"

        # Facebook commonly blocks server-side requests, so this branch is also
        # used for the user's pasted Marketplace text. Handle the labels Facebook
        # shows in the copyable listing panel and About this vehicle section.
        fb_loc = _first_match(text, [
            r"(?:listed\s+in|location)\s*[:\-]?\s*([A-Za-z][A-Za-z .'-]+,\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b",
            r"\b([A-Za-z][A-Za-z .'-]+,\s*(?:VIC|NSW|QLD|SA|WA|TAS|NT|ACT))\b",
        ], flags=re.I)
        if fb_loc:
            details["seller_location"] = re.sub(r"\s+", " ", fb_loc).strip()

        seller = _first_match(text, [
            r"(?:seller(?:\s+information)?|listed\s+by|seller\s+details)\s*[:\-]?\s*([A-Za-z][A-Za-z .'-]{1,60})",
        ], flags=re.I)
        if seller:
            details["seller_name"] = seller.strip(" ,-:")

        fb_odo = _first_match(text, [
            r"(?:driven|mileage|odometer|kilometres|kilometers)\s*[:\-]?\s*([\d,]+)\s*(?:km|kms|kilometres|kilometers)?\b",
            r"([\d,]{3,})\s*(?:km|kms)\s+(?:driven|mileage)\b",
        ], flags=re.I)
        if fb_odo:
            num = _number(fb_odo, integer=True)
            if num is not None and num < 10_000_000: details["odometer_km"] = num

        fb_trans = _first_match(text, [
            r"(automatic|manual|cvt|sports?\s+automatic)\s+transmission\b",
            r"transmission\s*[:\-]?\s*(sports?\s+automatic|automatic|manual|cvt)\b",
        ], flags=re.I)
        if fb_trans:
            t = fb_trans.lower()
            details["transmission"] = "Sports Automatic" if "sport" in t else ("CVT" if "cvt" in t else t.title())

        fb_colour = _first_match(text, [r"(?:exterior\s+colour|exterior\s+color|colour|color)\s*[:\-]?\s*([A-Za-z][A-Za-z -]{1,25})"], flags=re.I)
        if fb_colour:
            details["colour"] = re.split(r"\s+(?:interior|fuel|transmission|driven|mileage)\b", fb_colour, maxsplit=1, flags=re.I)[0].strip().title()

        fb_fuel = _first_match(text, [r"fuel(?:\s+type)?\s*[:\-]?\s*(diesel|petrol|gasoline|hybrid|electric|lpg)\b"], flags=re.I)
        if fb_fuel:
            f = fb_fuel.lower()
            details["fuel_type"] = "Petrol" if f == "gasoline" else f.title()

        fb_body = _first_match(text, [r"body\s+(?:type|style)\s*[:\-]?\s*([A-Za-z][A-Za-z ()/-]{1,30})"], flags=re.I)
        if fb_body:
            details["body_type"] = re.split(r"\s+(?:fuel|transmission|colour|color|driven|mileage)\b", fb_body, maxsplit=1, flags=re.I)[0].strip().title()

        fb_reg = _first_match(text, [r"(?:registration|rego)(?:\s+number)?\s*[:#\-]?\s*([A-Z0-9-]{2,10})\b"], flags=re.I)
        if fb_reg and fb_reg.lower() not in {"status", "expiry", "expires", "number"}:
            details["registration"] = fb_reg.upper()

        # Marketplace ads are fixed-price listings, not auctions.
        details.pop("auction_location", None)
        details.pop("auction_name", None)
        details.pop("lot_number", None)

    elif "carsales.com" in host:
        details["listing_source"] = "Carsales"
        if details.get("auction_location") and not details.get("seller_location"):
            details["seller_location"] = details.pop("auction_location")
        details.pop("auction_name", None)
        details.pop("lot_number", None)

    # Common cleanup: an auction/location field should never contain obvious site chrome.
    loc = details.get("auction_location") or ""
    if len(loc) > 100 or any(x in loc.lower() for x in ("category", "contact us", "privacy", "terms & conditions", "site links", "media complaints", "technical support")):
        details.pop("auction_location", None)

    return details

def _jsonld_blocks(page_html):
    objects = []
    text_blocks = []
    for block in re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        page_html or "",
        flags=re.I | re.S,
    ):
        cleaned = html.unescape(block).strip()
        try:
            obj = json.loads(cleaned)
            objects.append(obj)
            text_blocks.append(json.dumps(obj, ensure_ascii=False))
        except Exception:
            text_blocks.append(_strip_html(cleaned))
    return objects, text_blocks


def _collect_json_values(root, wanted_keys):
    """Iterative traversal avoids recursion failures on deeply nested site data."""
    wanted = {key.lower() for key in wanted_keys}
    found = []
    stack = [root]
    seen = 0
    while stack and seen < 50000:
        current = stack.pop()
        seen += 1
        if isinstance(current, dict):
            for key, value in current.items():
                if str(key).lower() in wanted:
                    found.append(value)
                if isinstance(value, (dict, list)):
                    stack.append(value)
        elif isinstance(current, list):
            for value in current:
                if isinstance(value, (dict, list)):
                    stack.append(value)
    return found


def _normalise_image_url(value):
    if not value:
        return ""
    if isinstance(value, dict):
        for key in ("url", "contentUrl", "thumbnailUrl"):
            if value.get(key):
                return _normalise_image_url(value.get(key))
        return ""
    value = html.unescape(str(value).strip()).replace("\\/", "/")
    if value.startswith("//"):
        value = "https:" + value
    if not value.startswith(("http://", "https://")):
        return ""
    return value


def _extract_photo_urls(page_html, jsonld_objects):
    candidates = []

    def add(value):
        if isinstance(value, list):
            for item in value:
                add(item)
            return
        url = _normalise_image_url(value)
        if not url:
            return
        lowered = url.lower()
        # Avoid common site chrome/logos where possible.
        if any(token in lowered for token in ("logo", "favicon", "sprite", "avatar", "icon-")):
            return
        if url not in candidates:
            candidates.append(url)

    for prop in ("og:image", "og:image:url", "twitter:image", "twitter:image:src"):
        add(_meta_content(page_html, prop))

    # Multiple metadata images.
    for match in re.findall(
        r'<meta[^>]+(?:property|name)=["\'](?:og:image|og:image:url|twitter:image|twitter:image:src)["\'][^>]+content=["\']([^"\']+)["\']',
        page_html or "",
        flags=re.I,
    ):
        add(match)

    for obj in jsonld_objects:
        for value in _collect_json_values(obj, ("image", "images", "photo", "photos", "thumbnailUrl", "contentUrl")):
            add(value)

    # Auction sites often embed image URLs in application state instead of JSON-LD.
    decoded = html.unescape(page_html or "").replace("\\/", "/")
    for match in re.findall(r'https?://[^\s"\'<>]+?\.(?:jpe?g|png|webp)(?:\?[^\s"\'<>]*)?', decoded, flags=re.I):
        add(match)
        if len(candidates) >= 20:
            break

    return candidates[:10]


def _auction_datetime_value(text, default_year=None):
    """Convert common Australian auction date text to datetime-local format."""
    if not text:
        return ""
    value = str(text).replace(".", ":")
    value = re.sub(r"\s+(?:AEST|AEDT|EST|EDT)\b", "", value, flags=re.I).strip(" ()")
    value = re.sub(r"^(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s*,?\s*", "", value, flags=re.I)
    value = re.sub(r"\s+", " ", value).strip(" ,")
    formats = (
        "%d %B %Y %H:%M", "%d %b %Y %H:%M", "%d %B %y %H:%M", "%d %b %y %H:%M",
        "%d %B %Y %I:%M %p", "%d %b %Y %I:%M %p", "%d %B %y %I:%M %p", "%d %b %y %I:%M %p",
        "%d/%m/%Y %H:%M", "%d/%m/%Y %I:%M %p", "%d-%m-%Y %H:%M", "%d-%m-%Y %I:%M %p",
    )
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).strftime("%Y-%m-%dT%H:%M")
        except ValueError:
            pass
    if default_year:
        for fmt in ("%d %B %H:%M", "%d %b %H:%M", "%d %B %I:%M %p", "%d %b %I:%M %p"):
            try:
                dt = datetime.strptime(value, fmt).replace(year=int(default_year))
                return dt.strftime("%Y-%m-%dT%H:%M")
            except ValueError:
                pass
    return ""


def _grays_sale_times(page_html, lot_url, headers):
    """Read Grays lot close time and, when available, its parent sale start time."""
    decoded = _strip_html(page_html or "")
    year = datetime.now().year
    finish_text = _first_match(decoded, [
        r"Closes\s*:\s*\(?\s*(\d{1,2}\s+[A-Za-z]+(?:\s+\d{2,4})?\s+\d{1,2}:\d{2}\s*(?:AEST|AEDT)?)",
        r"End\s+time\s*[:\-]?\s*(\d{1,2}\s+[A-Za-z]+(?:\s+\d{2,4})?\s+\d{1,2}[.:]\d{2}\s*(?:AM|PM)?\s*(?:AEST|AEDT)?)",
    ])
    finish = _auction_datetime_value((finish_text or "").replace(".", ":"), year)
    start = ""

    sale_match = re.search(r'href=["\']([^"\']*/sale/\d+/[^"\']*)["\']', page_html or "", flags=re.I)
    if sale_match:
        sale_url = urllib.parse.urljoin(lot_url, html.unescape(sale_match.group(1)).replace("\\/", "/"))
        try:
            _validate_public_http_url(sale_url)
            req = urllib.request.Request(sale_url, headers=headers)
            with urllib.request.urlopen(req, timeout=12) as response:
                sale_html = response.read(2_000_000).decode(response.headers.get_content_charset() or "utf-8", errors="ignore")
            sale_text = _strip_html(sale_html)
            start_text = _first_match(sale_text, [
                r"Start\s+time\s*[:\-]?\s*(\d{1,2}\s+[A-Za-z]+(?:\s+\d{2,4})?\s+\d{1,2}[.:]\d{2}\s*(?:AM|PM)?\s*(?:AEST|AEDT)?)",
            ])
            end_text = _first_match(sale_text, [
                r"End\s+time\s*[:\-]?\s*(\d{1,2}\s+[A-Za-z]+(?:\s+\d{2,4})?\s+\d{1,2}[.:]\d{2}\s*(?:AM|PM)?\s*(?:AEST|AEDT)?)",
            ])
            start = _auction_datetime_value((start_text or "").replace(".", ":"), year)
            finish = _auction_datetime_value((end_text or "").replace(".", ":"), year) or finish
        except Exception:
            pass
    return start, finish




def _embedded_label_windows(page_html, labels, before=180, after=1200):
    """Return readable text around important labels even when specs live in JS state.

    Some auction sites render their specification tables client-side. The values are
    still present in the downloaded HTML, but inside a large script block that normal
    visible-text extraction deliberately removes. Pulling small windows around known
    labels lets BAM read those values without treating the whole script as page text.
    """
    if not page_html:
        return ""
    raw = html.unescape(str(page_html))
    # Decode the common escaping used by React/Next/Salesforce state blobs.
    raw = (raw.replace("\\/", "/")
              .replace("\\u0026", "&")
              .replace("\\u003c", "<")
              .replace("\\u003e", ">")
              .replace("\\u0022", '"')
              .replace("\\n", " ")
              .replace("\\r", " ")
              .replace("\\t", " "))
    windows = []
    low = raw.lower()
    for label in labels:
        needle = str(label).lower()
        start = 0
        hits = 0
        while hits < 8:
            idx = low.find(needle, start)
            if idx < 0:
                break
            chunk = raw[max(0, idx-before): min(len(raw), idx+len(needle)+after)]
            # Keep script content, remove only markup/punctuation noise.
            chunk = re.sub(r"<[^>]+>", " ", chunk)
            chunk = re.sub(r"[{}\[\]\"']+", " ", chunk)
            # Client-side state often stores rows as label/value objects. Remove
            # the structural key names so the result reads like the visible table.
            chunk = re.sub(r"\b(?:label|displayValue|display_value|fieldName|field_name|value)\b\s*[:=]\s*", " ", chunk, flags=re.I)
            chunk = re.sub(r"\s+", " ", chunk).strip()
            if chunk:
                windows.append(chunk)
            hits += 1
            start = idx + len(needle)
    return " ".join(windows)


def _fetch_pickles_search_card(stock_id, request_headers):
    """Fetch Pickles' server-rendered search card for a stock number.

    Pickles' detail page currently renders many specs in browser JavaScript, while
    the search results expose odometer/seats/cylinders/engine/fuel/transmission/drive
    in normal HTML. This provides a reliable fallback for those fields.
    """
    stock_id = re.sub(r"\D", "", str(stock_id or ""))
    if not stock_id:
        return ""
    candidates = [
        f"https://www.pickles.com.au/used/search/lob/cars-motorcycles/cars?search={stock_id}",
        f"https://www.pickles.com.au/used/search/items?search={stock_id}",
    ]
    needle = f"Stock {stock_id}".lower()
    for candidate in candidates:
        try:
            req = urllib.request.Request(candidate, headers=request_headers)
            with urllib.request.urlopen(req, timeout=12) as response:
                ctype = (response.headers.get("Content-Type") or "").lower()
                if "html" not in ctype:
                    continue
                charset = response.headers.get_content_charset() or "utf-8"
                raw = response.read(4_000_000).decode(charset, errors="ignore")
            text = _strip_html(raw)
            idx = text.lower().find(needle)
            if idx >= 0:
                # The useful specs occur before the Stock marker on the result card.
                return text[max(0, idx-1600): min(len(text), idx+350)]
        except Exception:
            continue
    return ""


def _fetch_listing_page(url):
    parsed = _validate_public_http_url(url)
    source = _detect_listing_source(url)
    site_name = _listing_site_name(url)

    request_headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36",
        "Accept-Language": "en-AU,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Cache-Control": "no-cache",
    }
    req = urllib.request.Request(url, headers=request_headers)

    try:
        with urllib.request.urlopen(req, timeout=18) as response:
            ctype = (response.headers.get("Content-Type") or "").lower()
            if "text/html" not in ctype and "application/xhtml" not in ctype:
                raise ValueError("That link did not return a normal web listing page.")
            charset = response.headers.get_content_charset() or "utf-8"
            raw = response.read(4_000_000)
    except urllib.error.HTTPError as exc:
        if source == "Facebook Marketplace":
            raise ValueError("Facebook blocked automatic import. Copy the Marketplace listing title, price and description and use Import from Text.")
        raise ValueError(f"{site_name or 'The listing site'} blocked automatic import (HTTP {exc.code}). Copy the listing text and use Import from Text.")
    except urllib.error.URLError as exc:
        raise ValueError(f"BAM could not connect to that listing: {exc.reason}")
    except TimeoutError:
        raise ValueError("The listing site took too long to respond. Try again or use Import from Text.")

    page_html = raw.decode(charset, errors="ignore")
    title = _meta_content(page_html, "og:title") or _meta_content(page_html, "twitter:title")
    description = _meta_content(page_html, "og:description") or _meta_content(page_html, "description") or _meta_content(page_html, "twitter:description")
    if not title:
        match = re.search(r"<title[^>]*>(.*?)</title>", page_html, flags=re.I | re.S)
        title = _strip_html(match.group(1)) if match else ""

    jsonld_objects, jsonld_text = _jsonld_blocks(page_html)
    visible_text = _strip_html(page_html)

    # Pull small readable windows from the *entire* HTML around the labels used by
    # Pickles and Slattery. Their spec tables can be rendered from JavaScript state
    # well after the first few hundred KB of the document.
    label_windows = _embedded_label_windows(page_html, [
        "Odometer", "Odometer (Showing on)", "VIN", "Registration",
        "Registration Number", "Registration Expiry", "Registration Status",
        "Transmission", "Drive Type", "Body Type", "Fuel Type", "Fuel",
        "Engine Capacity", "Capacity", "Cylinders", "No of Seats",
        "No. of Seats", "Colour", "Color", "Trim", "Year Of Manufacture",
        "Sold Registered", "Sold on Consignment", "Item Details",
    ])

    # Pickles exposes the most useful specs in its server-rendered search result
    # card even when the detail page hides them behind client-side JavaScript.
    pickles_card = ""
    if "pickles.com" in (parsed.netloc or "").lower():
        stock_match = re.search(r"/(\d{6,12})(?:[/?#]|$)", urllib.parse.urlparse(url).path + "/")
        if stock_match:
            pickles_card = _fetch_pickles_search_card(stock_match.group(1), request_headers)

    # Keep a modest raw prefix as a generic fallback, then add targeted windows and
    # any Pickles card text. This avoids loading megabytes of unrelated script data.
    embedded_text = html.unescape(page_html).replace("\\/", "/")[:350000]
    combined = " ".join(part for part in (
        title, description, " ".join(jsonld_text), visible_text,
        label_windows, pickles_card, embedded_text
    ) if part)

    details = _extract_listing_details(combined, url=url, title=title, description=description)
    details["listing_url"] = url
    if source == "Auction":
        details["auction_url"] = url
        if site_name:
            details.setdefault("auction_name", site_name)
    details["photo_urls"] = _extract_photo_urls(page_html, jsonld_objects)

    if "grays.com" in (parsed.netloc or "").lower():
        auction_start, auction_finish = _grays_sale_times(page_html, url, request_headers)
        if auction_start:
            details["auction_start"] = auction_start
        if auction_finish:
            details["auction_finish"] = auction_finish

    # If Facebook returns a login/generic shell, do not pretend it imported.
    if source == "Facebook Marketplace":
        useful = any(details.get(key) for key in ("year", "make", "model", "asking_price", "odometer_km", "vin"))
        if not useful:
            raise ValueError("Facebook did not provide the actual Marketplace listing to BAM. Copy the listing title, price and description and use Import from Text.")

    return details


def _parse_imported_photo_urls(raw_value):
    if not raw_value:
        return []
    try:
        values = json.loads(raw_value)
    except (TypeError, json.JSONDecodeError):
        values = [x.strip() for x in str(raw_value).splitlines() if x.strip()]
    if not isinstance(values, list):
        return []
    result = []
    for value in values:
        url = _normalise_image_url(value)
        if url and url not in result:
            result.append(url)
    return result[:10]


def _download_listing_photo(url, referer=""):
    _validate_public_http_url(url)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/129.0 Safari/537.36",
        "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
    }
    if referer:
        headers["Referer"] = referer
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=15) as response:
            ctype = (response.headers.get("Content-Type") or "").split(";", 1)[0].lower().strip()
            if ctype not in {"image/jpeg", "image/jpg", "image/png", "image/webp"}:
                return None
            data = response.read(12 * 1024 * 1024 + 1)
            if not data or len(data) > 12 * 1024 * 1024:
                return None
    except Exception:
        return None

    ext_map = {"image/jpeg": "jpg", "image/jpg": "jpg", "image/png": "png", "image/webp": "webp"}
    ext = ext_map.get(ctype, "jpg")
    filename = f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}_listing.{ext}"
    (UPLOAD_DIR / filename).write_bytes(data)
    return filename


def _save_imported_listing_photos(conn, auction_vehicle_id, raw_urls, referer="", limit=10):
    saved = 0
    for url in _parse_imported_photo_urls(raw_urls):
        if saved >= limit:
            break
        filename = _download_listing_photo(url, referer=referer)
        if not filename:
            continue
        conn.execute("INSERT INTO auction_photos(auction_vehicle_id,filename,caption) VALUES(?,?,?)", (auction_vehicle_id, filename, "Imported from listing"))
        saved += 1
    return saved


# BAM Auction Watch - Version 25.3
# -----------------------------------------------------------------------------
AUCTION_TYPES = ("Car", "Boat", "Caravan", "Trailer", "Motorcycle", "Other")
AUCTION_STATUSES = ("Watching", "Contacted", "Negotiating", "Bidding", "Won", "Bought", "Lost", "Sold", "Passed In", "Removed")
BUYING_SOURCES = ("Auction", "Facebook Marketplace", "Gumtree", "Carsales", "Dealer", "Private Seller", "Other")
AUCTION_CONDITIONS = ("Excellent", "Good", "Average", "Poor", "Damaged", "BER", "Unknown")

# Version 25.3 - Quick Select catalog.
# These are suggestions only: Make and Model remain free-typing fields so an
# uncommon vehicle can still be entered without changing the program.
VEHICLE_MODEL_CATALOG = {
    "Toyota": ["86", "Camry", "Corolla", "C-HR", "Fortuner", "HiAce", "Hilux", "Kluger", "LandCruiser", "LandCruiser Prado", "Prius", "RAV4", "Yaris"],
    "Ford": ["Everest", "Falcon", "Focus", "Mustang", "Ranger", "Territory", "Transit"],
    "Holden": ["Astra", "Colorado", "Commodore", "Cruze", "Trax"],
    "Mazda": ["2", "3", "6", "BT-50", "CX-3", "CX-30", "CX-5", "CX-60", "CX-8", "CX-9", "MX-5"],
    "Nissan": ["370Z", "Dualis", "Juke", "Navara", "Patrol", "Pathfinder", "Qashqai", "X-Trail"],
    "Mitsubishi": ["ASX", "Eclipse Cross", "Lancer", "Outlander", "Pajero", "Pajero Sport", "Triton"],
    "Hyundai": ["Accent", "i30", "iLoad", "iMax", "Kona", "Santa Fe", "Sonata", "Tucson"],
    "Kia": ["Carnival", "Cerato", "Picanto", "Rio", "Sorento", "Sportage", "Stinger"],
    "Subaru": ["BRZ", "Forester", "Impreza", "Liberty", "Outback", "WRX", "XV"],
    "Volkswagen": ["Amarok", "Caddy", "Golf", "Passat", "Polo", "Tiguan", "Touareg", "Transporter"],
    "Isuzu": ["D-MAX", "MU-X"],
    "Suzuki": ["Baleno", "Grand Vitara", "Ignis", "Jimny", "Swift", "Vitara"],
    "Honda": ["Accord", "City", "Civic", "CR-V", "HR-V", "Jazz"],
    "Jeep": ["Cherokee", "Compass", "Gladiator", "Grand Cherokee", "Renegade", "Wrangler"],
    "Land Rover": ["Defender", "Discovery", "Discovery Sport", "Range Rover", "Range Rover Evoque", "Range Rover Sport"],
    "Lexus": ["ES", "GS", "IS", "LX", "NX", "RX", "UX"],
    "BMW": ["1 Series", "2 Series", "3 Series", "4 Series", "5 Series", "X1", "X3", "X5"],
    "Mercedes-Benz": ["A-Class", "C-Class", "E-Class", "GLA", "GLC", "GLE", "Sprinter", "Vito"],
    "Audi": ["A1", "A3", "A4", "A5", "Q2", "Q3", "Q5", "Q7"],
    "MG": ["MG3", "HS", "ZS"],
    "GWM": ["Cannon", "Ora", "Tank 300", "Tank 500"],
    "Haval": ["H2", "H6", "Jolion"],
    "LDV": ["D90", "Deliver 9", "G10", "T60"],
    "Ram": ["1500", "2500", "3500"],
    "Tesla": ["Model 3", "Model S", "Model X", "Model Y"],
    "Volvo": ["S60", "V60", "XC40", "XC60", "XC90"],
    "Skoda": ["Fabia", "Karoq", "Kodiaq", "Octavia", "Superb"],
    "Renault": ["Captur", "Kangoo", "Koleos", "Master", "Trafic"],
    "Peugeot": ["2008", "3008", "308", "5008", "Partner"],
}

# Version 25.4 - linked Variant suggestions.
# Variant remains free-typing so BAM can still accept any trim/grade not listed.
VEHICLE_VARIANT_CATALOG = {
    "Land Rover": {
        "Range Rover": ["Vogue", "Vogue SE", "HSE", "Autobiography", "Supercharged"],
        "Range Rover Sport": ["SE", "HSE", "HSE Dynamic", "Autobiography", "SVR"],
        "Range Rover Evoque": ["Pure", "Prestige", "Dynamic", "SE", "HSE"],
        "Discovery": ["S", "SE", "HSE", "HSE Luxury"],
        "Defender": ["S", "SE", "HSE", "X-Dynamic", "X"],
    },
    "Toyota": {
        "LandCruiser": ["GX", "GXL", "VX", "Sahara", "Sahara ZX", "GR Sport"],
        "LandCruiser Prado": ["GX", "GXL", "VX", "Kakadu"],
        "Hilux": ["WorkMate", "SR", "SR5", "Rogue", "GR Sport"],
        "RAV4": ["GX", "GXL", "Cruiser", "Edge"],
        "Camry": ["Ascent", "Ascent Sport", "SX", "SL"],
        "Corolla": ["Ascent Sport", "SX", "ZR"],
    },
    "Ford": {
        "Ranger": ["XL", "XLS", "XLT", "Sport", "Wildtrak", "Raptor"],
        "Everest": ["Ambiente", "Trend", "Sport", "Platinum"],
    },
    "Mazda": {
        "CX-5": ["Maxx", "Maxx Sport", "Touring", "GT", "Akera"],
        "BT-50": ["XS", "XT", "XTR", "GT", "SP", "Thunder"],
    },
    "Nissan": {
        "Patrol": ["Ti", "Ti-L", "Warrior"],
        "Navara": ["SL", "ST", "ST-X", "PRO-4X", "Warrior"],
        "X-Trail": ["ST", "ST-L", "Ti", "Ti-L"],
    },
    "Mitsubishi": {
        "Triton": ["GLX", "GLX+", "GLS", "GSR"],
        "Pajero Sport": ["GLX", "GLS", "Exceed", "GSR"],
        "Outlander": ["ES", "LS", "Aspire", "Exceed", "Exceed Tourer"],
    },
    "Isuzu": {
        "D-MAX": ["SX", "LS-M", "LS-U", "X-Terrain"],
        "MU-X": ["LS-M", "LS-U", "LS-T"],
    },
    "Volkswagen": {
        "Amarok": ["Core", "Life", "Style", "PanAmericana", "Aventura"],
        "Golf": ["Trendline", "Comfortline", "Highline", "GTI", "R"],
    },
    "Subaru": {
        "Forester": ["2.5i", "2.5i-L", "2.5i Premium", "2.5i-S", "Sport"],
        "Outback": ["AWD", "AWD Sport", "AWD Touring", "XT Sport", "XT Touring"],
    },
}


# Version 25.5 - Auction Watch make/model suggestions by asset type.
# Every field still accepts manual typing, so uncommon makes/models are supported.
AUCTION_ASSET_MODEL_CATALOG = {
    "Car": VEHICLE_MODEL_CATALOG,
    "Motorcycle": {
        "Honda": ["CB125E", "CB500F", "CBR500R", "CBR600RR", "CBR1000RR", "CRF250", "CRF300L", "CRF450R", "Gold Wing", "Rebel"],
        "Yamaha": ["MT-03", "MT-07", "MT-09", "MT-10", "R3", "R6", "R7", "R1", "Tenere 700", "WR450F"],
        "Kawasaki": ["Ninja 400", "Ninja 500", "Ninja 650", "Ninja ZX-6R", "Ninja ZX-10R", "Z400", "Z650", "Z900", "KLR650"],
        "Suzuki": ["GSX-R600", "GSX-R750", "GSX-R1000", "GSX-8R", "SV650", "V-Strom 650", "V-Strom 800", "DR-Z400"],
        "Harley-Davidson": ["Sportster", "Street Bob", "Fat Bob", "Fat Boy", "Low Rider", "Road King", "Street Glide"],
        "BMW": ["G 310", "F 750 GS", "F 850 GS", "R 1250 GS", "R 1300 GS", "S 1000 RR"],
        "KTM": ["390 Duke", "690 Enduro", "790 Duke", "890 Adventure", "1290 Super Adventure"],
        "Triumph": ["Bonneville", "Street Triple", "Speed Triple", "Tiger 900", "Tiger 1200"],
        "Ducati": ["Monster", "Panigale V2", "Panigale V4", "Multistrada", "Scrambler"],
    },
    "Caravan": {
        "Jayco": ["Journey", "Silverline", "Starcraft", "Expanda", "Discovery", "All-Terrain", "CrossTrak"],
        "New Age": ["Manta Ray", "Road Owl", "Desert Rose", "Big Red", "Wayfinder"],
        "Coromal": ["Element", "Princeton", "Lifestyle", "Magnum"],
        "Windsor": ["Genesis", "Rapid", "Statesman", "Silhouette"],
        "Avan": ["Aspire", "Infinity", "Frances", "Cruiseliner"],
        "Lotus": ["Freelander", "Trooper", "Off Grid", "Tremor"],
        "Zone RV": ["Sojourn", "Expedition", "Summit", "Peregrine"],
        "Kedron": ["Top Ender", "XC5", "Compact", "TE7"],
    },
    "Boat": {
        "Quintrex": ["Explorer", "Renegade", "Top Ender", "Fishabout", "Freestyler", "Trident"],
        "Stacer": ["Proline", "Sea Master", "Crossfire", "Ocean Ranger", "Wild Rider"],
        "Savage": ["Kestrel", "Scorpion", "Raptor", "Mako"],
        "Haines Hunter": ["V17L", "V19R", "SF535", "SF600", "675 Offshore"],
        "Cruise Craft": ["Explorer", "Outsider", "Resort", "F360"],
        "Bar Crusher": ["490", "535", "575", "615", "670", "730"],
        "Sea-Doo": ["Spark", "GTI", "GTR", "GTX", "RXP-X", "FishPro"],
        "Yamaha": ["FX", "VX", "WaveRunner", "AR195", "SX190"],
    },
    "Trailer": {
        "Custom": ["Box Trailer", "Car Trailer", "Plant Trailer", "Boat Trailer", "Enclosed Trailer", "Tipper Trailer"],
        "Mackay": ["Boat Trailer", "Car Trailer"],
        "Dunbier": ["Boat Trailer", "Jetski Trailer"],
        "Redco": ["Boat Trailer", "Box Trailer"],
    },
    "Other": {},
}

VEHICLE_MAKES = tuple(VEHICLE_MODEL_CATALOG.keys())
FUEL_TYPES = ("Petrol", "Diesel", "2 Stroke", "4 Stroke", "Hybrid", "Plug-in Hybrid", "Electric", "LPG", "Other")
TRANSMISSION_TYPES = ("Automatic", "Sports Automatic", "Manual", "CVT", "DCT", "Other")
DRIVE_TYPES = ("2WD", "4WD", "AWD", "FWD", "RWD", "Other")


def _auction_num(name, integer=False):
    raw = (request.form.get(name) or "").strip().replace(",", "").replace("$", "")
    if not raw:
        return None if integer else 0.0
    return int(float(raw)) if integer else float(raw)


def _auction_payload():
    make = (request.form.get("make") or "").strip()
    model = (request.form.get("model") or "").strip()
    if not make or not model:
        raise ValueError("Make and model are required.")
    asset_type = request.form.get("asset_type") or "Car"
    status = request.form.get("status") or "Watching"
    return {
        "status": status if status in AUCTION_STATUSES else "Watching",
        "asset_type": asset_type if asset_type in AUCTION_TYPES else "Other",
        "listing_source": (request.form.get("listing_source") or "Auction").strip() or "Auction",
        "seller_name": (request.form.get("seller_name") or "").strip() or None,
        "seller_phone": (request.form.get("seller_phone") or "").strip() or None,
        "seller_location": (request.form.get("seller_location") or "").strip() or None,
        "listing_url": (request.form.get("listing_url") or "").strip() or None,
        "date_first_seen": request.form.get("date_first_seen") or None,
        "last_checked": request.form.get("last_checked") or None,
        "asking_price": _auction_num("asking_price"),
        "negotiated_price": _auction_num("negotiated_price"),
        "auction_name": (request.form.get("auction_name") or "").strip() or None,
        "auction_location": (request.form.get("auction_location") or "").strip() or None,
        "auction_url": (request.form.get("auction_url") or "").strip() or None,
        "lot_number": (request.form.get("lot_number") or "").strip() or None,
        "auction_start": request.form.get("auction_start") or None,
        "auction_finish": request.form.get("auction_finish") or None,
        "year": _auction_num("year", True), "make": make, "model": model,
        "variant": (request.form.get("variant") or "").strip() or None,
        "vin": (request.form.get("vin") or "").strip().upper() or None,
        "registration": (request.form.get("registration") or "").strip().upper() or None,
        "registration_status": (request.form.get("registration_status") or "").strip() or None,
        "reserve_status": (request.form.get("reserve_status") or "Unknown").strip() or "Unknown",
        "body_type": (request.form.get("body_type") or "").strip() or None,
        "seat_count": _auction_num("seat_count", True) if (request.form.get("seat_count") or "").strip() else None,
        "odometer_km": _auction_num("odometer_km", True),
        "engine_hours": _auction_num("engine_hours") if (request.form.get("engine_hours") or "").strip() else None,
        "engine_size": (request.form.get("engine_size") or "").strip() or None,
        "engine_cc": _auction_num("engine_cc", True) if (request.form.get("engine_cc") or "").strip() else None,
        "engine_cylinders": (request.form.get("engine_cylinders") or "").strip() or None,
        "length_m": _auction_num("length_m") if (request.form.get("length_m") or "").strip() else None,
        "berths": _auction_num("berths", True) if (request.form.get("berths") or "").strip() else None,
        "axles": _auction_num("axles", True) if (request.form.get("axles") or "").strip() else None,
        "tare_weight_kg": _auction_num("tare_weight_kg") if (request.form.get("tare_weight_kg") or "").strip() else None,
        "atm_kg": _auction_num("atm_kg") if (request.form.get("atm_kg") or "").strip() else None,
        "gtm_kg": _auction_num("gtm_kg") if (request.form.get("gtm_kg") or "").strip() else None,
        "ball_weight_kg": _auction_num("ball_weight_kg") if (request.form.get("ball_weight_kg") or "").strip() else None,
        "width_m": _auction_num("width_m") if (request.form.get("width_m") or "").strip() else None,
        "height_m": _auction_num("height_m") if (request.form.get("height_m") or "").strip() else None,
        "caravan_features": (request.form.get("caravan_features") or "").strip() or None,
        "boat_type": (request.form.get("boat_type") or "").strip() or None,
        "hull_material": (request.form.get("hull_material") or "").strip() or None,
        "engine_make": (request.form.get("engine_make") or "").strip() or None,
        "engine_model": (request.form.get("engine_model") or "").strip() or None,
        "horsepower": _auction_num("horsepower") if (request.form.get("horsepower") or "").strip() else None,
        "trailer_included": 1 if request.form.get("trailer_included") else 0,
        "trailer_registration": (request.form.get("trailer_registration") or "").strip().upper() or None,
        "capacity_people": _auction_num("capacity_people", True) if (request.form.get("capacity_people") or "").strip() else None,
        "boat_features": (request.form.get("boat_features") or "").strip() or None,
        "trailer_features": (request.form.get("trailer_features") or "").strip() or None,
        "colour": (request.form.get("colour") or "").strip() or None,
        "interior": (request.form.get("interior") or "").strip() or None,
        "transmission": (request.form.get("transmission") or "").strip() or None,
        "drive_type": (request.form.get("drive_type") or "").strip() or None,
        "fuel_type": (request.form.get("fuel_type") or "").strip() or None,
        "tow_bar": 1 if request.form.get("tow_bar") else 0,
        "condition_grade": (request.form.get("condition_grade") or "Unknown").strip(),
        "condition_notes": (request.form.get("condition_notes") or "").strip() or None,
        "current_bid": _auction_num("current_bid"), "max_bid": _auction_num("max_bid"),
        "sold_price": _auction_num("sold_price"), "auction_fees": _auction_num("auction_fees"),
        "transport_cost": _auction_num("transport_cost"), "other_costs": _auction_num("other_costs"),
        "quick_sale_value": _auction_num("quick_sale_value"), "repair_allowance": _auction_num("repair_allowance"),
        "target_profit": _auction_num("target_profit"), "rego_ppsr_cost": _auction_num("rego_ppsr_cost"),
        "boat_engine_cost": _auction_num("boat_engine_cost"), "boat_hull_cost": _auction_num("boat_hull_cost"),
        "boat_trailer_cost": _auction_num("boat_trailer_cost"),
        "comparable_price_1": _auction_num("comparable_price_1"),
        "comparable_price_2": _auction_num("comparable_price_2"),
        "comparable_price_3": _auction_num("comparable_price_3"),
        "comparable_price_4": _auction_num("comparable_price_4"),
        "comparable_price_5": _auction_num("comparable_price_5"),
    }


def _auction_market_value(conn, item):
    rows = conn.execute("""
        SELECT sold_price FROM auction_vehicles
        WHERE sold_price>0 AND LOWER(make)=LOWER(?) AND LOWER(model)=LOWER(?) AND id<>?
        ORDER BY COALESCE(auction_finish,created_at) DESC LIMIT 30
    """, (item["make"], item["model"], item["id"])).fetchall()
    auction_prices = sorted(float(r["sold_price"]) for r in rows if float(r["sold_price"] or 0)>0)
    retail = conn.execute("""
        SELECT s.sale_price_inc_gst AS price FROM sales s JOIN vehicles v ON v.id=s.vehicle_id
        WHERE s.sale_price_inc_gst>0 AND LOWER(v.make)=LOWER(?) AND LOWER(v.model)=LOWER(?)
        ORDER BY s.sale_date DESC LIMIT 20
    """, (item["make"], item["model"])).fetchall()
    retail_prices = sorted(float(r["price"]) for r in retail if float(r["price"] or 0)>0)
    if auction_prices:
        mid = statistics.median(auction_prices)
        low = auction_prices[max(0, int((len(auction_prices)-1)*.25))]
        high = auction_prices[min(len(auction_prices)-1, int((len(auction_prices)-1)*.75))]
        source = "BAM Auction Watch sold history"
        if retail_prices:
            high = max(high, statistics.median(retail_prices)); source += " + BAM retail sales"
    elif retail_prices:
        retail_mid = statistics.median(retail_prices)
        low, mid, high = retail_mid*.78, retail_mid*.86, retail_mid
        source = "BAM retail sales converted to an auction estimate"
    else:
        base = float(item["max_bid"] or item["current_bid"] or 0)
        low, mid, high = base, base*1.10 if base else 0, base*1.20 if base else 0
        source = "Entered bid only - no comparable BAM history yet"
    return round(low,2), round(mid,2), round(high,2), source, len(auction_prices), len(retail_prices)


AUCTION_PAGE = r"""
<!doctype html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>BAM Buying Watch</title>
<style>body{font-family:Arial;background:#0f172a;color:#e5e7eb;margin:0}.wrap{max-width:1450px;margin:auto;padding:24px}.top{display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap}.brand{font-size:30px;font-weight:800}.sub,.muted{color:#94a3b8}.btn{display:inline-block;padding:10px 14px;border-radius:9px;background:#2563eb;color:white;text-decoration:none;border:0;font-weight:700;cursor:pointer}.secondary{background:#334155}.panel,.card{background:#111827;border:1px solid #334155;border-radius:14px}.panel{padding:16px;margin-top:16px}.filters{display:grid;grid-template-columns:2fr repeat(5,1fr) auto;gap:10px}.filters input,.filters select{padding:10px;border-radius:8px;border:1px solid #475569;background:#0b1220;color:white}.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:14px;margin-top:16px}.card{overflow:hidden}.thumb{height:190px;background:#020617;display:flex;align-items:center;justify-content:center;color:#64748b}.thumb img{width:100%;height:100%;object-fit:cover}.cardbody{padding:14px}.title{font-size:20px;font-weight:800}.pill{display:inline-block;background:#1e293b;border:1px solid #475569;border-radius:999px;padding:4px 8px;margin:3px 2px;font-size:12px}.price{font-size:18px;font-weight:800;margin-top:8px}.actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}.flash{background:#1e3a8a;padding:10px;border-radius:8px;margin:10px 0}@media(max-width:900px){.filters{grid-template-columns:1fr 1fr}}@media(max-width:550px){.filters{grid-template-columns:1fr}}</style>
</head><body><div class='wrap'><div class='top'><div><div class='brand'>🔎 BAM Buying Watch</div><div class='sub'>Cars • Boats • Caravans • Trailers • Motorcycles • Other</div></div><div><a class='btn secondary' href='{{url_for("dashboard")}}'>← BAM Dashboard</a> <a class='btn' href='{{url_for("auction_add")}}'>+ Add Watch Vehicle</a></div></div>
{% with messages=get_flashed_messages(with_categories=true) %}{% for cat,msg in messages %}<div class='flash'>{{msg}}</div>{% endfor %}{% endwith %}
<div class='panel'><form class='filters' method='get'><input name='q' value='{{q}}' placeholder='Search make, model, source, seller, auction, location...'><select name='asset_type'><option value=''>All types</option>{% for x in types %}<option {{'selected' if asset_type==x else ''}}>{{x}}</option>{% endfor %}</select><select name='source'><option value=''>All sources</option>{% for x in sources %}<option {{'selected' if source_filter==x else ''}}>{{x}}</option>{% endfor %}</select><select name='status'><option value=''>All status</option>{% for x in statuses %}<option {{'selected' if status==x else ''}}>{{x}}</option>{% endfor %}</select><input name='make' value='{{make}}' placeholder='Make'><input name='model' value='{{model}}' placeholder='Model'><button class='btn'>Search</button></form></div>
<div class='cards'>{% for v in rows %}<div class='card'><div class='thumb'>{% if v.thumbnail %}<img src='{{url_for("uploaded_file",filename=v.thumbnail)}}'>{% else %}No photo yet{% endif %}</div><div class='cardbody'><div class='title'>{{v.year or ''}} {{v.make}} {{v.model}}</div><div class='muted'>{{v.variant or ''}} • {{v.listing_source or 'Auction'}}{% if (v.listing_source or 'Auction')=='Auction' %} • {{v.auction_name or 'Auction not set'}} • Lot {{v.lot_number or '-'}}{% else %} • {{v.seller_location or 'Location not set'}}{% endif %}</div><div><span class='pill'>{{v.asset_type}}</span><span class='pill'>{{v.status}}</span>{% if v.odometer_km %}<span class='pill'>{{'{:,}'.format(v.odometer_km)}} km</span>{% endif %}{% if v.transmission %}<span class='pill'>{{v.transmission}}</span>{% endif %}{% if v.drive_type %}<span class='pill'>{{v.drive_type}}</span>{% endif %}</div><div class='price'>{% if (v.listing_source or 'Auction')=='Auction' %}Current ${{'{:,.0f}'.format(v.current_bid or 0)}} · Sold ${{'{:,.0f}'.format(v.sold_price or 0)}}{% else %}Asking ${{'{:,.0f}'.format(v.asking_price or 0)}} · Negotiated ${{'{:,.0f}'.format(v.negotiated_price or 0)}}{% endif %}</div><div class='muted'>Finishes: {{v.auction_finish or 'Not set'}} · {{v.colour or 'Colour not set'}} · {{v.condition_grade or 'Unknown'}}</div><div class='actions'><a class='btn' href='{{url_for("auction_detail",auction_id=v.id)}}'>Open</a><a class='btn secondary' href='{{url_for("auction_watch",make=v.make,model=v.model)}}'>Same Model History</a></div></div></div>{% else %}<div class='panel'>No Buying Watch vehicles match these filters yet.</div>{% endfor %}</div></div></body></html>
"""

AUCTION_FORM = r"""
<!doctype html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Buying Watch Vehicle</title><style>body{font-family:Arial;background:#0f172a;color:#e5e7eb;margin:0}.wrap{max-width:1150px;margin:auto;padding:24px}.panel{background:#111827;border:1px solid #334155;border-radius:14px;padding:18px;margin-top:14px}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.full{grid-column:1/-1}label{display:block;font-size:12px;color:#94a3b8;margin-bottom:4px}input,select,textarea{width:100%;box-sizing:border-box;padding:10px;border-radius:8px;border:1px solid #475569;background:#0b1220;color:#fff}textarea{min-height:100px}.btn{padding:10px 14px;border:0;border-radius:9px;background:#2563eb;color:white;text-decoration:none;font-weight:700;cursor:pointer}.secondary{background:#334155}.danger{background:#b91c1c}.good{background:#15803d}.top,.actions{display:flex;gap:8px;justify-content:space-between;align-items:center;flex-wrap:wrap}.actions{justify-content:flex-start;margin-top:14px}.photos{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}.photos img{width:100%;height:120px;object-fit:cover;border-radius:8px}.photos img{cursor:zoom-in}.bam-lightbox{display:none;position:fixed;inset:0;z-index:99999;background:rgba(0,0,0,.94);align-items:center;justify-content:center;padding:30px}.bam-lightbox.open{display:flex}.bam-lightbox img{max-width:94vw;max-height:90vh;width:auto;height:auto;object-fit:contain;border-radius:8px;box-shadow:0 10px 45px rgba(0,0,0,.55)}.bam-lightbox-close,.bam-lightbox-prev,.bam-lightbox-next{position:fixed;border:0;background:rgba(15,23,42,.78);color:#fff;cursor:pointer;border-radius:10px;font-size:34px;line-height:1;padding:10px 14px}.bam-lightbox-close{top:16px;right:18px}.bam-lightbox-prev{left:18px;top:50%;transform:translateY(-50%)}.bam-lightbox-next{right:18px;top:50%;transform:translateY(-50%)}.notice{background:#1e3a8a;padding:10px;border-radius:8px;margin:10px 0}.value{font-size:24px;font-weight:800}.asset-field.hidden,.auction-field.hidden,.market-field.hidden,.gumtree-description-field.hidden{display:none}.valuation-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}.valuation-card{background:#0b1220;border:1px solid #334155;border-radius:10px;padding:12px}.valuation-card span{display:block;color:#94a3b8;font-size:12px;margin-bottom:6px}.valuation-card strong{font-size:18px}.buycalc{border:2px solid #2563eb}.deal-score{font-size:25px;font-weight:900}.deal-good{color:#4ade80}.deal-caution{color:#fbbf24}.deal-stop{color:#f87171}.calc-note{color:#94a3b8;font-size:12px;margin-top:8px}@media(max-width:900px){.valuation-grid{grid-template-columns:repeat(2,1fr)}}@media(max-width:800px){.grid{grid-template-columns:1fr 1fr}.photos{grid-template-columns:repeat(2,1fr)}}@media(max-width:520px){.grid{grid-template-columns:1fr}}</style></head><body><div class='wrap'><div class='top'><h1>{{'Edit' if item else 'Add'}} Watch Vehicle</h1><a class='btn secondary' href='{{url_for("auction_watch")}}'>← Buying Watch</a></div>{% with messages=get_flashed_messages(with_categories=true) %}{% for cat,msg in messages %}<div class='notice'>{{msg}}</div>{% endfor %}{% endwith %}<div class='panel'><h2>🔗 Import Listing Details</h2><div class='grid'>
<div class='full'><label>Facebook Marketplace / Auction / Carsales / Gumtree Link</label><div style='display:flex;gap:8px'><input id='import_url' placeholder='Paste the listing link here'><button type='button' class='btn good' id='import_link_btn' style='width:auto;white-space:nowrap'>Import from Link</button></div></div>
<div class='full'><label>Facebook / Gumtree Listing Text <span class='muted'>(not the link)</span></label><textarea id='import_text' placeholder='Open the ad and copy the actual listing text here — title, price, kilometres, location, vehicle details and description. Keep the link in the box above.'></textarea><div class='muted' style='margin-top:6px'>Facebook and Gumtree can block automatic link reading. BAM will keep the link above and use this pasted text to fill the vehicle.</div><button type='button' class='btn secondary' id='import_text_btn' style='margin-top:8px'>Import Listing Text</button></div>
<div class='full'><div id='import_status' class='notice' style='display:none'></div></div>
</div></div>
<div class='panel'><form method='post' enctype='multipart/form-data'><input type='hidden' id='imported_photo_urls' name='imported_photo_urls' value=''><div class='grid'>
<div><label>Vehicle Type</label><select id='auction_asset_type' name='asset_type'>{% for x in types %}<option {{'selected' if item and item.asset_type==x else ''}}>{{x}}</option>{% endfor %}</select></div><div><label>Source</label><select id='listing_source' name='listing_source'>{% for x in sources %}<option {{'selected' if item and item.listing_source==x else ''}}>{{x}}</option>{% endfor %}</select></div><div><label>Status</label><select name='status'>{% for x in statuses %}<option {{'selected' if item and item.status==x else ''}}>{{x}}</option>{% endfor %}</select></div><div class='market-field'><label>Seller Name</label><input name='seller_name' value='{{item.seller_name or "" if item else ""}}'></div><div class='market-field'><label>Seller Phone</label><input name='seller_phone' value='{{item.seller_phone or "" if item else ""}}'></div><div class='market-field'><label>Seller / Listing Location</label><input name='seller_location' value='{{item.seller_location or "" if item else ""}}'></div><div class='market-field full'><label>Listing URL</label><input id='listing_url' name='listing_url' placeholder='Facebook Marketplace, Carsales, Gumtree or other link' value='{{item.listing_url or "" if item else ""}}'></div><div class='market-field'><label>Date First Seen</label><input type='date' name='date_first_seen' value='{{item.date_first_seen or "" if item else ""}}'></div><div class='market-field'><label>Last Checked</label><input type='date' name='last_checked' value='{{item.last_checked or "" if item else ""}}'></div><div></div><div class='auction-field'><label>Lot Number</label><input id='lot_number' name='lot_number' value='{{item.lot_number or "" if item else ""}}'></div>
<div class='auction-field'><label>Auction Name</label><input name='auction_name' value='{{item.auction_name or "" if item else ""}}'></div><div class='auction-field'><label>Auction Location</label><input name='auction_location' value='{{item.auction_location or "" if item else ""}}'></div><div class='auction-field'><label>Auction Web Link</label><div style='display:flex;gap:8px'><input id='auction_url' name='auction_url' placeholder='Paste auction webpage link' value='{{item.auction_url or "" if item else ""}}'><button type='button' class='btn secondary' id='open_auction_url_btn' style='width:auto;white-space:nowrap'>🌐 Open Webpage</button></div></div>
<div class='auction-field'><label>Auction Starts</label><input type='datetime-local' name='auction_start' value='{{item.auction_start or "" if item else ""}}'></div><div class='auction-field'><label>Auction Finishes</label><input type='datetime-local' name='auction_finish' value='{{item.auction_finish or "" if item else ""}}'></div><div></div>
<div><label>Year</label><input type='number' id='year' name='year' value='{{item.year or "" if item else ""}}'></div><div><label>Make *</label><input id='auction_make' name='make' list='make_options' autocomplete='off' required value='{{item.make or "" if item else ""}}'><datalist id='make_options'></datalist></div><div><label>Model *</label><input id='auction_model' name='model' list='model_options' autocomplete='off' required value='{{item.model or "" if item else ""}}'><datalist id='model_options'></datalist></div><div><label>Variant</label><input id='auction_variant' name='variant' list='variant_options' autocomplete='off' value='{{item.variant or "" if item else ""}}'><datalist id='variant_options'></datalist></div><div><label>VIN / Chassis</label><input id='vin' name='vin' value='{{item.vin or "" if item else ""}}'></div><div><label>Registration</label><input id='registration' name='registration' value='{{item.registration or "" if item else ""}}'></div><div><label>Registration Status</label><input id='registration_status' name='registration_status' list='registration_status_options' autocomplete='off' value='{{item.registration_status or "" if item else ""}}' placeholder='e.g. Sold Registered, Sold on Consignment'><datalist id='registration_status_options'><option value='Sold Registered, Sold on Consignment'><option value='Sold Registered'><option value='Sold Unregistered'><option value='Registered'><option value='Unregistered'></datalist></div>
<div class='asset-field car'><label>Body Type</label><input id='body_type' name='body_type' list='body_type_options' value='{{item.body_type or "" if item else ""}}'><datalist id='body_type_options'><option value='SUV'><option value='Sedan'><option value='Wagon'><option value='Hatchback'><option value='Ute'><option value='Van'><option value='Coupe'><option value='Convertible'><option value='Cab Chassis'></datalist></div><div class='asset-field car'><label>No. of Seats</label><input type='number' min='1' max='99' id='seat_count' name='seat_count' value='{{item.seat_count or "" if item else ""}}'></div>
<div class='asset-field car motorcycle'><label>Kilometres</label><input type='number' id='odometer_km' name='odometer_km' value='{{item.odometer_km or "" if item else ""}}'></div><div class='asset-field boat'><label>Engine Hours</label><input type='number' step='0.1' name='engine_hours' value='{{item.engine_hours or "" if item else ""}}'></div><div><label>Colour</label><input id='colour' name='colour' value='{{item.colour or "" if item else ""}}'></div>
<div class='asset-field car boat'><label>Engine Size</label><input id='engine_size' name='engine_size' list='engine_size_options' placeholder='e.g. 3.0L, 4.4L' value='{{item.engine_size or "" if item else ""}}'><datalist id='engine_size_options'><option value='1.0L'><option value='1.2L'><option value='1.5L'><option value='1.6L'><option value='1.8L'><option value='2.0L'><option value='2.2L'><option value='2.4L'><option value='2.5L'><option value='2.8L'><option value='3.0L'><option value='3.2L'><option value='3.5L'><option value='4.0L'><option value='4.4L'><option value='4.5L'><option value='4.6L'><option value='5.0L'><option value='5.7L'><option value='6.2L'></datalist></div>
<div class='asset-field motorcycle'><label>Engine Size (cc)</label><input type='number' id='engine_cc' name='engine_cc' placeholder='e.g. 650' value='{{item.engine_cc or "" if item else ""}}'></div>
<div class='asset-field car motorcycle'><label>Cylinders</label><input name='engine_cylinders' list='cylinder_options' value='{{item.engine_cylinders or "" if item else ""}}'><datalist id='cylinder_options'><option value='1'><option value='2'><option value='3'><option value='4'><option value='5'><option value='6'><option value='8'><option value='10'><option value='12'></datalist></div>
<div class='asset-field car'><label>Interior</label><select name='interior'><option></option>{% for x in ['Cloth','Leather','Vinyl','Other'] %}<option {{'selected' if item and item.interior==x else ''}}>{{x}}</option>{% endfor %}</select></div><div class='asset-field car motorcycle'><label>Transmission</label><select id='transmission' name='transmission'><option></option>{% for x in transmissions %}<option {{'selected' if item and item.transmission==x else ''}}>{{x}}</option>{% endfor %}</select></div><div class='asset-field car'><label>Drive</label><select id='drive_type' name='drive_type'><option></option>{% for x in drives %}<option {{'selected' if item and item.drive_type==x else ''}}>{{x}}</option>{% endfor %}</select></div>
<div class='asset-field car boat motorcycle'><label>Fuel Type</label><select id='fuel_type' name='fuel_type'><option></option>{% for x in fuels %}<option {{'selected' if item and item.fuel_type==x else ''}}>{{x}}</option>{% endfor %}</select></div><div><label>Condition</label><select name='condition_grade'>{% for x in conditions %}<option {{'selected' if (item and item.condition_grade==x) or (not item and x=='Unknown') else ''}}>{{x}}</option>{% endfor %}</select></div><div class='asset-field car' style='padding-top:24px'><label><input type='checkbox' name='tow_bar' style='width:auto' {{'checked' if item and item.tow_bar else ''}}> Tow bar fitted</label></div>
<div class='asset-field caravan trailer boat'><label>Length (metres)</label><input type='number' step='0.01' name='length_m' value='{{item.length_m or "" if item else ""}}'></div>
<div class='asset-field caravan'><label>Berths</label><input type='number' name='berths' value='{{item.berths or "" if item else ""}}'></div>
<div class='asset-field caravan trailer'><label>Axles</label><input type='number' name='axles' value='{{item.axles or "" if item else ""}}'></div>
<div class='asset-field caravan trailer'><label>Tare Weight (kg)</label><input type='number' step='0.1' name='tare_weight_kg' value='{{item.tare_weight_kg or "" if item else ""}}'></div>
<div class='asset-field caravan trailer'><label>ATM (kg)</label><input type='number' step='0.1' name='atm_kg' value='{{item.atm_kg or "" if item else ""}}'></div>
<div class='asset-field caravan trailer'><label>GTM (kg)</label><input type='number' step='0.1' name='gtm_kg' value='{{item.gtm_kg or "" if item else ""}}'></div>
<div class='asset-field caravan trailer'><label>Ball Weight (kg)</label><input type='number' step='0.1' name='ball_weight_kg' value='{{item.ball_weight_kg or "" if item else ""}}'></div>
<div class='asset-field caravan trailer'><label>Width (metres)</label><input type='number' step='0.01' name='width_m' value='{{item.width_m or "" if item else ""}}'></div>
<div class='asset-field caravan trailer'><label>Height (metres)</label><input type='number' step='0.01' name='height_m' value='{{item.height_m or "" if item else ""}}'></div>
<div class='asset-field caravan full'><label>Caravan Features</label><textarea name='caravan_features' placeholder='Ensuite, air conditioning, solar, batteries, fridge, suspension...'>{{item.caravan_features or "" if item else ""}}</textarea></div>
<div class='asset-field boat'><label>Boat Type</label><select name='boat_type'><option></option>{% for x in ['Fishing','Runabout','Bowrider','Cabin','Centre Console','PWC / Jet Ski','Ski / Wake','Sailing','Other'] %}<option {{'selected' if item and item.boat_type==x else ''}}>{{x}}</option>{% endfor %}</select></div>
<div class='asset-field boat'><label>Hull Material</label><select name='hull_material'><option></option>{% for x in ['Aluminium','Fibreglass','Steel','Wood','Inflatable','Other'] %}<option {{'selected' if item and item.hull_material==x else ''}}>{{x}}</option>{% endfor %}</select></div>
<div class='asset-field boat'><label>Engine Make</label><input name='engine_make' placeholder='e.g. Yamaha' value='{{item.engine_make or "" if item else ""}}'></div>
<div class='asset-field boat'><label>Engine Model</label><input name='engine_model' value='{{item.engine_model or "" if item else ""}}'></div>
<div class='asset-field boat'><label>Horsepower</label><input type='number' step='0.1' name='horsepower' value='{{item.horsepower or "" if item else ""}}'></div>
<div class='asset-field boat'><label>Capacity (people)</label><input type='number' name='capacity_people' value='{{item.capacity_people or "" if item else ""}}'></div>
<div class='asset-field boat'><label>Trailer Registration</label><input name='trailer_registration' value='{{item.trailer_registration or "" if item else ""}}'></div>
<div class='asset-field boat' style='padding-top:24px'><label><input type='checkbox' name='trailer_included' style='width:auto' {{'checked' if item and item.trailer_included else ''}}> Trailer included</label></div>
<div class='asset-field boat full'><label>Boat Features</label><textarea name='boat_features' rows='16' style='min-height:260px' placeholder='Sounder, GPS, canopy, electric anchor, safety gear...'>{{item.boat_features or "" if item else ""}}</textarea></div>
<div class='asset-field trailer full'><label>Trailer Features</label><textarea name='trailer_features' placeholder='Brakes, dimensions, ramps, winch, cage, tipping...'>{{item.trailer_features or "" if item else ""}}</textarea></div>
<div class='market-field'><label>Asking Price</label><input type='number' step='.01' id='asking_price' name='asking_price' value='{{item.asking_price or 0 if item else 0}}'></div><div class='market-field'><label>Negotiated Price</label><input type='number' step='.01' name='negotiated_price' value='{{item.negotiated_price or 0 if item else 0}}'></div><div class='auction-field'><label>Reserve Status</label><select id='reserve_status' name='reserve_status'><option {{'selected' if not item or not item.reserve_status or item.reserve_status=='Unknown' else ''}}>Unknown</option><option {{'selected' if item and item.reserve_status=='No Reserve' else ''}}>No Reserve</option><option {{'selected' if item and item.reserve_status=='Reserve' else ''}}>Reserve</option></select></div><div class='auction-field'><label>Current Bid</label><input type='number' step='.01' id='current_bid' name='current_bid' value='{{item.current_bid or 0 if item else 0}}'></div><div class='auction-field'><label>My Maximum Bid</label><input type='number' step='.01' name='max_bid' value='{{item.max_bid or 0 if item else 0}}'></div><div class='auction-field'><label>Sold Price</label><input type='number' step='.01' name='sold_price' value='{{item.sold_price or 0 if item else 0}}'></div><div class='auction-field'><label>Auction Fees</label><input type='number' step='.01' name='auction_fees' value='{{item.auction_fees or 0 if item else 0}}'></div><div><label>Transport Cost</label><input type='number' step='.01' name='transport_cost' value='{{item.transport_cost or 0 if item else 0}}'></div><div><label>Other Costs</label><input type='number' step='.01' name='other_costs' value='{{item.other_costs or 0 if item else 0}}'></div><div><label>Quick-Sale Value</label><input type='number' step='.01' name='quick_sale_value' value='{{item.quick_sale_value or 0 if item else 0}}'></div><div><label>Repairs / Reconditioning</label><input type='number' step='.01' name='repair_allowance' value='{{item.repair_allowance or 0 if item else 0}}'></div><div><label>Registration / PPSR / Transfer</label><input type='number' step='.01' name='rego_ppsr_cost' value='{{item.rego_ppsr_cost or 0 if item else 0}}'></div><div><label>Target Profit</label><input type='number' step='.01' name='target_profit' value='{{item.target_profit or 0 if item else 0}}'></div><div class='asset-field boat'><label>Boat Engine Service / Repairs</label><input type='number' step='.01' name='boat_engine_cost' value='{{item.boat_engine_cost or 0 if item else 0}}'></div><div class='asset-field boat'><label>Hull Repairs</label><input type='number' step='.01' name='boat_hull_cost' value='{{item.boat_hull_cost or 0 if item else 0}}'></div><div class='asset-field boat'><label>Trailer / Trailer Rego</label><input type='number' step='.01' name='boat_trailer_cost' value='{{item.boat_trailer_cost or 0 if item else 0}}'></div>
<div class='full' id='pre_save_valuation'><div class='panel' style='margin:8px 0 0'><h2>📊 Enter Market Valuation Before Saving</h2><div class='muted'>Use BAM Automatic Market Valuation to research similar Australian Cars, Boats, Caravans or Trailers and fill up to five advertised prices automatically. You can still edit them manually. BAM calculates the market ranges, puts Market Mid into Quick-Sale Value and updates the buying guide before you save.</div><div class='grid' style='margin-top:12px'><div><label>Comparable 1 ($)</label><input class='pre-comp' type='number' step='.01' min='0' name='comparable_price_1' value='{{item.comparable_price_1 or "" if item else ""}}'></div><div><label>Comparable 2 ($)</label><input class='pre-comp' type='number' step='.01' min='0' name='comparable_price_2' value='{{item.comparable_price_2 or "" if item else ""}}'></div><div><label>Comparable 3 ($)</label><input class='pre-comp' type='number' step='.01' min='0' name='comparable_price_3' value='{{item.comparable_price_3 or "" if item else ""}}'></div><div><label>Comparable 4 ($)</label><input class='pre-comp' type='number' step='.01' min='0' name='comparable_price_4' value='{{item.comparable_price_4 or "" if item else ""}}'></div><div><label>Comparable 5 ($)</label><input class='pre-comp' type='number' step='.01' min='0' name='comparable_price_5' value='{{item.comparable_price_5 or "" if item else ""}}'></div></div><div class='valuation-grid' style='margin-top:12px'><div class='valuation-card'><span>Market Low</span><strong id='pre_market_low'>$0</strong></div><div class='valuation-card'><span>Market Mid / Quick-Sale</span><strong id='pre_market_mid'>$0</strong></div><div class='valuation-card'><span>Market High</span><strong id='pre_market_high'>$0</strong></div><div class='valuation-card'><span>Wholesale Range</span><strong id='pre_wholesale'>$0 – $0</strong></div><div class='valuation-card'><span>Trade Range</span><strong id='pre_trade'>$0 – $0</strong></div><div class='valuation-card'><span>Private Range</span><strong id='pre_private'>$0 – $0</strong></div></div><div class='actions'><button type='button' class='btn good' id='pre_calc_valuation_btn'>Calculate Market Valuation & Buying Guide</button></div><div class='valuation-grid' style='margin-top:12px'><div class='valuation-card'><span>Target Buy Price</span><strong id='pre_target_buy'>$0</strong></div><div class='valuation-card'><span>Maximum Recommended Bid</span><strong id='pre_max_bid'>$0</strong></div><div class='valuation-card'><span>Walk-Away Price</span><strong id='pre_walkaway'>$0</strong></div><div class='valuation-card'><span>Expected Profit</span><strong id='pre_expected_profit'>$0</strong></div><div class='valuation-card'><span>Deal Score</span><strong id='pre_deal_score'>-</strong></div></div><div id='pre_deal_message' class='notice' style='margin-top:12px'>Enter comparable prices and BAM will calculate the buying guide.</div></div></div><div class='full'><label>Listing Description <span class='muted'>(Facebook, Gumtree, auction or other listing)</span></label><textarea id='listing_description' placeholder='Paste or add the full listing description here. BAM will copy it into Condition / Inspection Notes below.'></textarea><div class='muted' style='margin-top:6px'>Use this for Cars, Boats, Caravans and Trailers. The description is kept in Condition / Inspection Notes so the original seller or auction information stays with the watch item.</div></div>
<div class='full'><label>Condition / Inspection Notes</label><textarea id='condition_notes' name='condition_notes'>{{item.condition_notes or "" if item else ""}}</textarea></div><div class='full'><label>Add Listing Photos (maximum 10 total)</label><input id='listing_photo_files' type='file' name='photos' accept='image/*' multiple><div class='actions' style='margin-top:8px'><button type='button' class='btn secondary' id='paste_photo_btn'>📋 Paste Copied Photo</button></div><div class='muted' style='margin-top:6px'>Facebook / Gumtree: right-click a listing photo and choose Copy image, then click Paste Copied Photo. Repeat for more photos, or use Choose Files to select several at once.</div><div id='local_photo_preview' class='photos' style='margin-top:10px'></div><div id='imported_photo_preview' class='photos' style='margin-top:10px'></div></div></div><div class='actions'><button class='btn'>Save Watch Vehicle</button>{% if item %}<a class='btn secondary' href='{{url_for("auction_value",auction_id=item.id)}}'>Get Valuation</a>{% endif %}</div></form></div>
{% if not item %}<div class='panel'><h2>🇦🇺 BAM Market Valuation</h2><div class='muted'>BAM Automatic Market Valuation researches the Australian market and fills the comparable prices and buying guide before you save. Google is kept as a manual cross-check.</div><div class='actions'><button type='button' class='btn good' id='bam_auto_valuation_btn'>✨ BAM Automatic Market Valuation</button><button type='button' class='btn secondary' id='google_market_valuation_btn'>🔎 Open Google Cross-Check ↗</button></div></div>{% endif %}
{% if item %}<div class='panel buycalc'><h2>💰 BAM Auction Buy Calculator & Deal Score</h2><div class='valuation-grid'><div class='valuation-card'><span>Quick-Sale Value</span><strong id='calc_quick'>$0</strong></div><div class='valuation-card'><span>Total Costs</span><strong id='calc_costs'>$0</strong></div><div class='valuation-card'><span>🎯 Target Buy Price</span><strong id='calc_target'>$0</strong></div><div class='valuation-card'><span>🟠 Maximum Recommended Bid</span><strong id='calc_max'>$0</strong></div><div class='valuation-card'><span>🔴 Walk-Away Price</span><strong id='calc_walkaway'>$0</strong></div><div class='valuation-card'><span>Expected Profit at Current Price</span><strong id='calc_profit'>$0</strong></div><div class='valuation-card'><span>Deal Score</span><strong id='calc_score' class='deal-score'>—</strong></div></div><div id='calc_message' class='notice' style='margin-top:12px'>Enter a Quick-Sale Value and BAM will calculate the deal.</div><div class='calc-note'>Target Buy uses the low end of BAM wholesale value. Maximum Recommended Bid uses the high end of BAM wholesale value. Walk-Away is the absolute ceiling from Quick-Sale Value less all costs and target profit. Boats also include engine, hull and trailer allowances.</div><div class='actions'><button type='button' class='btn good' id='calculate_bid_btn'>Calculate My Maximum Bid</button><button type='button' class='btn secondary' id='use_market_quick_btn'>Use BAM Market Mid</button></div></div><div class='panel'><h2>🇦🇺 BAM Valuation Hub</h2><div class='valuation-grid'><div class='valuation-card'><span>Private Sale</span><strong>${{'{:,.0f}'.format(item.private_value_low or 0)}} – ${{'{:,.0f}'.format(item.private_value_high or 0)}}</strong></div><div class='valuation-card'><span>Wholesale</span><strong>${{'{:,.0f}'.format(item.wholesale_value_low or 0)}} – ${{'{:,.0f}'.format(item.wholesale_value_high or 0)}}</strong></div><div class='valuation-card'><span>Trade-In</span><strong>${{'{:,.0f}'.format(item.trade_value_low or 0)}} – ${{'{:,.0f}'.format(item.trade_value_high or 0)}}</strong></div><div class='valuation-card'><span>Dealer Retail</span><strong>${{'{:,.0f}'.format(item.dealer_value_low or 0)}} – ${{'{:,.0f}'.format(item.dealer_value_high or 0)}}</strong></div><div class='valuation-card'><span>Suggested Buy / Max Bid</span><strong>${{'{:,.0f}'.format(item.suggested_buy_price or 0)}}</strong></div></div><div style='margin-top:12px'><b>Provider:</b> {{item.valuation_provider or 'BAM internal history'}} &nbsp; <b>Confidence:</b> {{item.valuation_confidence or 'Not calculated'}}</div><div class='muted' style='margin-top:6px'>{{item.valuation_source or 'Click Get Valuation for BAM internal pricing, or use the Carsales buttons for a live Australian market check.'}}</div><div class='actions'><form method='post' action='{{url_for("auction_automatic_valuation",auction_id=item.id)}}' style='display:inline'><button class='btn good'>✨ BAM Automatic Market Valuation</button></form><button type='button' class='btn secondary' id='google_market_valuation_btn'>🔎 Open Google Cross-Check ↗</button><a class='btn good' target='_blank' rel='noopener' href='{{carsales_valuation_url}}'>Carsales Free Valuation ↗</a><a class='btn secondary' target='_blank' rel='noopener' href='{{carsales_search_url}}'>Carsales Comparable Search ↗</a></div><div class='muted' style='margin-top:8px'>Google Market Valuation uses the details currently entered above and builds an Australian comparison search tailored to Cars, Boats, Caravans and Trailers before you set BAM valuation figures.</div><form method='post' action='{{url_for("auction_comparable_value",auction_id=item.id)}}' style='margin-top:16px'><h3 style='margin-bottom:8px'>📊 Comparable Market Prices</h3><div class='muted' style='margin-bottom:10px'>Enter up to five advertised prices you find for similar cars, boats, caravans or trailers. BAM will calculate indicative private, wholesale, trade-in and dealer-retail ranges from those real comparables.</div><div class='grid'><div><label>Comparable 1 ($)</label><input type='number' step='.01' min='0' name='comparable_price_1' value='{{item.comparable_price_1 or ""}}'></div><div><label>Comparable 2 ($)</label><input type='number' step='.01' min='0' name='comparable_price_2' value='{{item.comparable_price_2 or ""}}'></div><div><label>Comparable 3 ($)</label><input type='number' step='.01' min='0' name='comparable_price_3' value='{{item.comparable_price_3 or ""}}'></div><div><label>Comparable 4 ($)</label><input type='number' step='.01' min='0' name='comparable_price_4' value='{{item.comparable_price_4 or ""}}'></div><div><label>Comparable 5 ($)</label><input type='number' step='.01' min='0' name='comparable_price_5' value='{{item.comparable_price_5 or ""}}'></div></div><div class='actions'><button class='btn good'>Calculate Market Valuation</button></div></form></div><div class='panel'><h2>Photos ({{photos|length}} / 10)</h2><div class='photos'>{% for p in photos %}<div><img src='{{url_for("uploaded_file",filename=p.filename)}}'><form method='post' action='{{url_for("auction_delete_photo",auction_id=item.id,photo_id=p.id)}}'><button class='btn danger' style='margin-top:5px'>Delete</button></form></div>{% else %}<div>No photos yet.</div>{% endfor %}</div></div><div class='panel'><h2>Bought / won this vehicle?</h2>{% if item.won_vehicle_id %}<div class='notice'>Already transferred to BAM Vehicle Stock.</div><a class='btn' href='{{url_for("vehicle_detail",vehicle_id=item.won_vehicle_id)}}'>Open Vehicle Stock Record</a>{% else %}<form method='post' action='{{url_for("auction_transfer",auction_id=item.id)}}'><div class='grid'><div><label>Purchased By / Ownership</label><select name='sale_ownership'><option>BAM Joint</option><option>Barry</option><option>Matt</option></select></div><div><label>Purchase Date</label><input type='date' name='purchase_date' value='{{today}}'></div></div><div class='actions'><button class='btn good'>✓ Add to BAM Vehicle Stock</button></div></form>{% endif %}</div><div class='panel'><form method='post' action='{{url_for("auction_delete",auction_id=item.id)}}' onsubmit='return confirm("Delete this auction vehicle?")'><button class='btn danger'>Delete Watch Vehicle</button></form></div>{% endif %}</div><div id='bam_lightbox' class='bam-lightbox' aria-hidden='true'><button type='button' class='bam-lightbox-close' aria-label='Close'>×</button><button type='button' class='bam-lightbox-prev' aria-label='Previous photo'>‹</button><img id='bam_lightbox_image' alt='Large photo'><button type='button' class='bam-lightbox-next' aria-label='Next photo'>›</button></div><script>
const BAM_MODELS = {{ model_catalog_json|safe }};
const BAM_VARIANTS = {{ variant_catalog_json|safe }};
const BAM_ASSET_MODELS = {{ asset_model_catalog_json|safe }};
const assetTypeInput = document.getElementById('auction_asset_type');
const makeInput = document.getElementById('auction_make');
const modelInput = document.getElementById('auction_model');
const variantInput = document.getElementById('auction_variant');
const makeList = document.getElementById('make_options');
const modelList = document.getElementById('model_options');
const variantList = document.getElementById('variant_options');
const googleMarketBtn = document.getElementById('google_market_valuation_btn');
const bamAutoValuationBtn = document.getElementById('bam_auto_valuation_btn');
if(bamAutoValuationBtn){
  bamAutoValuationBtn.addEventListener('click',async()=>{
    const original=bamAutoValuationBtn.textContent;
    bamAutoValuationBtn.disabled=true; bamAutoValuationBtn.textContent='⏳ Researching Australian market...';
    try{
      const details={}; document.querySelectorAll('[name]').forEach(el=>{ if(el.name && el.type!=='file' && el.type!=='password') details[el.name]=el.value; });
      const r=await fetch('{{url_for("auction_automatic_valuation_preview")}}',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(details)});
      const data=await r.json(); if(!r.ok || !data.ok) throw new Error(data.error||'Valuation failed');
      const comps=document.querySelectorAll('.pre-comp'); (data.prices||[]).slice(0,5).forEach((v,i)=>{if(comps[i]) comps[i].value=v;});
      const quick=document.querySelector('[name="quick_sale_value"]'); if(quick) quick.value=data.market_mid||0;
      const calc=document.getElementById('pre_calc_valuation_btn'); if(calc) calc.click();
      alert('BAM found '+(data.prices||[]).length+' current Australian comparable listing(s). The valuation and buying guide have been filled in. Save the watch item to store them.');
    }catch(err){ alert('BAM Automatic Market Valuation: '+err.message); }
    finally{ bamAutoValuationBtn.disabled=false; bamAutoValuationBtn.textContent=original; }
  });
}
const openAuctionUrlBtn = document.getElementById('open_auction_url_btn');
if(openAuctionUrlBtn){
  openAuctionUrlBtn.addEventListener('click',()=>{
    const el=document.getElementById('auction_url');
    let url=(el && el.value ? el.value.trim() : '');
    if(!url){ alert('Enter the auction webpage link first.'); return; }
    if(!/^https?:\/\//i.test(url)) url='https://'+url;
    window.open(url,'_blank','noopener,noreferrer');
  });
}
if(googleMarketBtn){
  googleMarketBtn.addEventListener('click',()=>{
    const value=(name)=>{const el=document.querySelector(`[name="${name}"]`); return el ? (el.value||'').trim() : '';};
    const asset=value('asset_type') || 'Car';
    const parts=[value('year'),value('make'),value('model'),value('variant')];
    const listingIdentity=value('auction_name') || value('caravan_features') || value('trailer_features') || value('boat_features');
    if(!value('make') && !value('model') && listingIdentity) parts.push(listingIdentity.substring(0,180));
    if(asset==='Boat'){
      parts.push(value('boat_type'),value('hull_material'),value('engine_make'),value('engine_model'));
      if(value('horsepower')) parts.push(value('horsepower')+'hp');
      if(value('engine_hours')) parts.push(value('engine_hours')+' engine hours');
      if(value('length_m')) parts.push(value('length_m')+'m');
      parts.push('boat for sale Australia','Boatsonline Gumtree Facebook Marketplace dealer private');
    }else if(asset==='Caravan'){
      if(value('length_m')) parts.push(value('length_m')+'m');
      if(value('berths')) parts.push(value('berths')+' berth');
      if(value('tare_weight_kg')) parts.push(value('tare_weight_kg')+'kg tare');
      parts.push('caravan for sale Australia','Caravancampingsales Gumtree Facebook Marketplace dealer private');
    }else if(asset==='Trailer'){
      if(value('length_m')) parts.push(value('length_m')+'m');
      if(value('tare_weight_kg')) parts.push(value('tare_weight_kg')+'kg tare');
      if(value('atm_kg')) parts.push(value('atm_kg')+'kg ATM');
      parts.push('trailer for sale Australia','Gumtree Facebook Marketplace dealer private');
    }else{
      if(value('odometer_km')) parts.push(value('odometer_km')+' km');
      parts.push(asset.toLowerCase(),'for sale Australia','Carsales Gumtree Facebook Marketplace dealer private');
    }
    const query=parts.filter(Boolean).join(' ');
    if(!value('make') && !value('model') && !listingIdentity){alert('Enter the make/model or import the listing details first.');return;}
    window.open('https://www.google.com/search?q='+encodeURIComponent(query),'_blank','noopener');
  });
}

function catalogueKey(obj, value){
  const typed=(value||'').trim().toLowerCase();
  return Object.keys(obj||{}).find(k=>k.toLowerCase()===typed);
}
function refreshVariants(){
  if(!makeInput || !modelInput || !variantList) return;
  const makeKey=catalogueKey(BAM_VARIANTS, makeInput.value);
  const models=makeKey ? BAM_VARIANTS[makeKey] : {};
  const modelKey=catalogueKey(models, modelInput.value);
  variantList.innerHTML='';
  (modelKey ? models[modelKey] : []).forEach(v=>{const o=document.createElement('option');o.value=v;variantList.appendChild(o);});
}
function currentAssetCatalog(){
  const type=(assetTypeInput && assetTypeInput.value) || 'Car';
  return BAM_ASSET_MODELS[type] || {};
}
function refreshMakes(){
  if(!makeList) return;
  const catalog=currentAssetCatalog();
  makeList.innerHTML='';
  Object.keys(catalog).forEach(m=>{const o=document.createElement('option');o.value=m;makeList.appendChild(o);});
  refreshModels();
}
function refreshModels(){
  if(!makeInput || !modelList) return;
  const catalog=currentAssetCatalog();
  const key=catalogueKey(catalog, makeInput.value);
  modelList.innerHTML='';
  (key ? catalog[key] : []).forEach(m=>{const o=document.createElement('option');o.value=m;modelList.appendChild(o);});
  refreshVariants();
}
function refreshSourceFields(){
  const sourceInput=document.getElementById('listing_source');
  const source=(sourceInput && sourceInput.value) || 'Auction';
  const isAuction=source==='Auction';
  document.querySelectorAll('.auction-field').forEach(el=>el.classList.toggle('hidden', !isAuction));
  document.querySelectorAll('.market-field').forEach(el=>el.classList.toggle('hidden', isAuction));
  document.querySelectorAll('.gumtree-description-field').forEach(el=>el.classList.toggle('hidden', source!=='Gumtree'));
}
function refreshAssetFields(){
  const type=((assetTypeInput && assetTypeInput.value) || 'Car').toLowerCase();
  document.querySelectorAll('.asset-field').forEach(el=>{
    const show=el.classList.contains(type);
    el.classList.toggle('hidden', !show);
  });
  refreshMakes();
  refreshSourceFields();
}
if(assetTypeInput){
  assetTypeInput.addEventListener('change', refreshAssetFields);
}
const listingSourceInput=document.getElementById('listing_source');
if(listingSourceInput){listingSourceInput.addEventListener('change', refreshSourceFields);}
if(makeInput){
  makeInput.addEventListener('input', refreshModels);
  makeInput.addEventListener('change', refreshModels);
}
if(modelInput){
  modelInput.addEventListener('input', refreshVariants);
  modelInput.addEventListener('change', refreshVariants);
}

function setImportedField(name, value){
  if(value === undefined || value === null || value === '') return;
  const el=document.getElementById(name) || document.querySelector(`[name="${name}"]`);
  if(!el) return;
  el.value=value;
  el.dispatchEvent(new Event('input',{bubbles:true}));
  el.dispatchEvent(new Event('change',{bubbles:true}));
}
function showImportStatus(message, good=false){
  const box=document.getElementById('import_status');
  if(!box) return;
  box.style.display='block';
  box.textContent=message;
  box.style.background=good ? '#14532d' : '#7f1d1d';
}
function renderImportedPhotos(urls){
  const hidden=document.getElementById('imported_photo_urls');
  const preview=document.getElementById('imported_photo_preview');
  const list=Array.isArray(urls) ? urls.slice(0,10) : [];
  if(hidden) hidden.value=JSON.stringify(list);
  if(!preview) return;
  preview.innerHTML='';
  list.forEach(url=>{
    const img=document.createElement('img');
    img.src=url;
    img.alt='Imported listing photo';
    img.referrerPolicy='no-referrer';
    preview.appendChild(img);
  });
}
async function importListing(payload){
  showImportStatus('Reading listing…', true);
  try{
    const response=await fetch('{{url_for("auction_import_listing")}}',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify(payload)
    });
    const data=await response.json();
    if(!data.ok){
      showImportStatus(data.error || 'Could not import listing.');
      return;
    }
    const d=data.details || {};
    [
      'listing_source','listing_url','seller_name','seller_phone','seller_location',
      'auction_name','auction_location','auction_url','lot_number','auction_start','auction_finish',
      'year','make','model','variant','vin','registration','registration_status','body_type','seat_count','reserve_status','odometer_km','engine_size','engine_cc',
      'engine_cylinders','colour','interior','transmission','drive_type','fuel_type',
      'asking_price','current_bid','sold_price','status','condition_grade','condition_notes',
      'asset_type','length_m','width_m','height_m','berths','axles','tare_weight_kg','atm_kg','gtm_kg','ball_weight_kg','caravan_features',
      'engine_hours','boat_type','hull_material','engine_make','engine_model','horsepower','capacity_people','trailer_registration','boat_features','trailer_features'
    ].forEach(k=>setImportedField(k,d[k]));
    const trailerIncluded=document.getElementById('trailer_included') || document.querySelector('[name="trailer_included"]');
    if(trailerIncluded && d.trailer_included !== undefined && d.trailer_included !== null){
      trailerIncluded.checked=Boolean(Number(d.trailer_included) || d.trailer_included === true);
      trailerIncluded.dispatchEvent(new Event('change',{bubbles:true}));
    }
    const importedDescription=document.getElementById('listing_description');
    if(importedDescription && d.condition_notes){ importedDescription.value=d.condition_notes; }
    renderImportedPhotos(d.photo_urls || []);
    refreshAssetFields();
    refreshModels();
    refreshVariants();
    const count=(d.photo_urls || []).length;
    showImportStatus(`Details imported${count ? ` with ${count} photo${count===1?'':'s'}` : ''}. Please check them before saving.`, true);
  }catch(e){
    showImportStatus('Could not import listing. Paste the listing text instead.');
  }
}
function marketplaceSourceFromUrl(url){
  const u=(url||'').toLowerCase();
  if(u.includes('facebook.com') || u.includes('fb.com')) return 'Facebook Marketplace';
  if(u.includes('gumtree.com')) return 'Gumtree';
  if(u.includes('carsales.com')) return 'Carsales';
  return '';
}
function rememberMarketplaceLink(url){
  const source=marketplaceSourceFromUrl(url);
  if(source){
    setImportedField('listing_source',source);
    setImportedField('listing_url',url);
    refreshSourceFields();
  }
  return source;
}
const importUrlInput=document.getElementById('import_url');
if(importUrlInput){
  importUrlInput.addEventListener('input',()=>rememberMarketplaceLink((importUrlInput.value||'').trim()));
  importUrlInput.addEventListener('paste',()=>setTimeout(()=>rememberMarketplaceLink((importUrlInput.value||'').trim()),0));
}
const importLinkBtn=document.getElementById('import_link_btn');
if(importLinkBtn){
  importLinkBtn.addEventListener('click',()=>{
    const url=(document.getElementById('import_url').value||'').trim();
    if(!url){ showImportStatus('Paste a listing link first.'); return; }
    rememberMarketplaceLink(url);
    importListing({url:url});
  });
}
const listingDescription=document.getElementById('listing_description');
const conditionNotes=document.getElementById('condition_notes');
if(listingDescription && conditionNotes){
  listingDescription.addEventListener('input',()=>{ conditionNotes.value=listingDescription.value; });
  listingDescription.addEventListener('paste',()=>setTimeout(()=>{ conditionNotes.value=listingDescription.value; },0));
}

const preCalcBtn=document.getElementById('pre_calc_valuation_btn');
function money(v){return '$'+Math.round(v||0).toLocaleString('en-AU');}
function calculatePreSaveValuation(){
  const prices=[...document.querySelectorAll('.pre-comp')].map(x=>Number(x.value||0)).filter(x=>x>0).sort((a,b)=>a-b);
  if(!prices.length){alert('Enter at least one comparable advertised price first.');return;}
  const mid=prices.length%2 ? prices[(prices.length-1)/2] : (prices[prices.length/2-1]+prices[prices.length/2])/2;
  const low=prices[0], high=prices[prices.length-1];
  const set=(id,text)=>{const e=document.getElementById(id);if(e)e.textContent=text;};
  set('pre_market_low',money(low)); set('pre_market_mid',money(mid)); set('pre_market_high',money(high));
  set('pre_wholesale',money(low*.62)+' – '+money(mid*.72));
  set('pre_trade',money(low*.68)+' – '+money(mid*.78));
  set('pre_private',money(low*.92)+' – '+money(high*.97));
  const quick=document.querySelector('[name="quick_sale_value"]'); if(quick){quick.value=mid.toFixed(2);quick.dispatchEvent(new Event('input',{bubbles:true}));}
  const n=(name)=>{const e=document.querySelector(`[name="${name}"]`);const v=parseFloat(e&&e.value);return Number.isFinite(v)?v:0;};
  const asset=((document.querySelector('[name="asset_type"]')||{}).value||'Car').toLowerCase();
  let costs=n('auction_fees')+n('transport_cost')+n('repair_allowance')+n('rego_ppsr_cost')+n('other_costs'); if(asset==='boat') costs+=n('boat_engine_cost')+n('boat_hull_cost')+n('boat_trailer_cost');
  const tp=n('target_profit'), targetBuy=Math.max(0,low*.62-costs-tp), maxBid=Math.max(0,mid*.72-costs-tp), walk=Math.max(0,mid-costs-tp);
  const src=((document.querySelector('[name="listing_source"]')||{}).value||'Auction'); const current=src==='Auction'?n('current_bid'):(n('negotiated_price')||n('asking_price')); const profit=mid-current-costs;
  let score='READY',msg='Target '+money(targetBuy)+' | Maximum '+money(maxBid)+' | Walk away above '+money(walk)+'.'; if(current>0){if(current<=targetBuy){score='STRONG BUY';msg='Current price is at or below BAM Target Buy.';}else if(current<=maxBid){score='GOOD BUY';msg='Current price is within BAM recommended buying range.';}else if(current<=walk){score='CAUTION';msg='Above BAM recommended bid. Walk-Away is '+money(walk)+'.';}else{score='DO NOT BID';msg='Current price is above BAM Walk-Away Price.';}}
  set('pre_target_buy',money(targetBuy));set('pre_max_bid',money(maxBid));set('pre_walkaway',money(walk));set('pre_expected_profit',(profit<0?'-':'')+money(Math.abs(profit)));set('pre_deal_score',score);set('pre_deal_message',msg);
}
if(preCalcBtn) preCalcBtn.addEventListener('click',calculatePreSaveValuation);
const importTextBtn=document.getElementById('import_text_btn');
if(importTextBtn){
  importTextBtn.addEventListener('click',()=>{
    const pasted=(document.getElementById('import_text').value||'').trim();
    const url=(document.getElementById('import_url').value||'').trim();
    const source=rememberMarketplaceLink(url);
    if(!pasted){ showImportStatus('Paste the actual advertisement text into the lower box first.'); return; }
    if(/^https?:\/\//i.test(pasted)){
      showImportStatus('The lower box needs the advertisement text, not another link. Keep the link in the top box, then copy the title, price, kilometres, location and description from the ad.');
      return;
    }
    importListing({text:pasted,url:'',source_url:url,source_hint:source});
  });
}


const listingPhotoFiles=document.getElementById('listing_photo_files');
const localPhotoPreview=document.getElementById('local_photo_preview');
let copiedListingPhotos=[];
function refreshLocalPhotoPreview(){
  if(!localPhotoPreview) return;
  localPhotoPreview.innerHTML='';
  const files=listingPhotoFiles ? Array.from(listingPhotoFiles.files || []) : [];
  files.slice(0,10).forEach(file=>{
    const img=document.createElement('img');
    img.src=URL.createObjectURL(file);
    img.alt='Listing photo ready to save';
    img.onload=()=>URL.revokeObjectURL(img.src);
    localPhotoPreview.appendChild(img);
  });
}
function rebuildListingPhotoInput(){
  if(!listingPhotoFiles) return;
  const dt=new DataTransfer();
  copiedListingPhotos.slice(0,10).forEach(file=>dt.items.add(file));
  listingPhotoFiles.files=dt.files;
  refreshLocalPhotoPreview();
}
if(listingPhotoFiles){
  listingPhotoFiles.addEventListener('change',()=>{
    copiedListingPhotos=Array.from(listingPhotoFiles.files || []).slice(0,10);
    rebuildListingPhotoInput();
  });
}
const pastePhotoBtn=document.getElementById('paste_photo_btn');
if(pastePhotoBtn){
  pastePhotoBtn.addEventListener('click',async()=>{
    if(copiedListingPhotos.length>=10){ showImportStatus('Maximum 10 photos reached.'); return; }
    if(!navigator.clipboard || !navigator.clipboard.read){
      showImportStatus('This browser cannot paste copied images here. Use Choose Files instead.');
      return;
    }
    try{
      const items=await navigator.clipboard.read();
      let added=0;
      for(const item of items){
        const type=item.types.find(t=>t.startsWith('image/'));
        if(!type || copiedListingPhotos.length>=10) continue;
        const blob=await item.getType(type);
        const ext=(type.split('/')[1] || 'png').replace('jpeg','jpg');
        copiedListingPhotos.push(new File([blob],`listing-photo-${Date.now()}-${added+1}.${ext}`,{type:type}));
        added++;
      }
      if(!added){ showImportStatus('No copied photo was found. Right-click the listing photo, choose Copy image, then try again.'); return; }
      rebuildListingPhotoInput();
      showImportStatus(`${added} copied photo${added===1?'':'s'} added. ${copiedListingPhotos.length}/10 ready to save.`,true);
    }catch(e){
      showImportStatus('The browser could not read the copied photo. Allow clipboard access, or use Choose Files instead.');
    }
  });
}


// v25.13.2 - large photo viewer for Buying Watch, imported and pasted photos.
(function(){
  const box=document.getElementById('bam_lightbox');
  const big=document.getElementById('bam_lightbox_image');
  if(!box || !big) return;
  let photos=[], index=0;
  function collect(){
    photos=Array.from(document.querySelectorAll('.photos img')).filter(img=>img.src && !img.closest('#bam_lightbox'));
  }
  function show(i){
    collect(); if(!photos.length) return;
    index=(i+photos.length)%photos.length;
    big.src=photos[index].currentSrc || photos[index].src;
    big.alt=photos[index].alt || 'Large photo';
    box.classList.add('open'); box.setAttribute('aria-hidden','false');
    document.body.style.overflow='hidden';
  }
  function close(){box.classList.remove('open');box.setAttribute('aria-hidden','true');big.removeAttribute('src');document.body.style.overflow='';}
  document.addEventListener('click',e=>{
    const img=e.target.closest('.photos img');
    if(img && !img.closest('#bam_lightbox')){collect();show(photos.indexOf(img));return;}
    if(e.target===box || e.target.closest('.bam-lightbox-close')) close();
    if(e.target.closest('.bam-lightbox-prev')) show(index-1);
    if(e.target.closest('.bam-lightbox-next')) show(index+1);
  });
  document.addEventListener('keydown',e=>{
    if(!box.classList.contains('open')) return;
    if(e.key==='Escape') close();
    if(e.key==='ArrowLeft') show(index-1);
    if(e.key==='ArrowRight') show(index+1);
  });
})();

// v25.14.1 - live Auction Buy Calculator / Deal Score.
(function(){
  const byName=name=>document.querySelector(`[name="${name}"]`);
  const num=name=>{const el=byName(name);const n=parseFloat(el && el.value);return Number.isFinite(n)?n:0;};
  const money=n=>'$'+Math.max(0,n).toLocaleString('en-AU',{maximumFractionDigits:0});
  const quickEl=byName('quick_sale_value');
  const calcBtn=document.getElementById('calculate_bid_btn');
  const marketBtn=document.getElementById('use_market_quick_btn');
  const marketMid={{ (item.market_mid or 0) if item else 0 }};
  const wholesaleLow={{ (item.wholesale_value_low or 0) if item else 0 }};
  const wholesaleHigh={{ (item.wholesale_value_high or 0) if item else 0 }};

  function calculate(){
    if(!quickEl) return;
    const quick=num('quick_sale_value');
    const isBoat=((assetTypeInput && assetTypeInput.value)||'').toLowerCase()==='boat';
    const costs=num('auction_fees')+num('transport_cost')+num('repair_allowance')+num('rego_ppsr_cost')+num('other_costs')
      +(isBoat ? num('boat_engine_cost')+num('boat_hull_cost')+num('boat_trailer_cost') : 0);
    const target=num('target_profit');
    const targetBuy=Math.max(0,(wholesaleLow>0 ? wholesaleLow : quick*.62)-costs-target);
    const maxBid=Math.max(0,(wholesaleHigh>0 ? wholesaleHigh : quick*.72)-costs-target);
    const walkAway=Math.max(0,quick-costs-target);
    const source=((document.getElementById('listing_source')||{}).value||'Auction');
    let current=source==='Auction' ? num('current_bid') : (num('negotiated_price') || num('asking_price'));
    const profit=quick-current-costs;
    let score='—', message='Enter a Quick-Sale Value and BAM will calculate the deal.';
    if(quick>0){
      if(current<=0){score='READY';message='Target '+money(targetBuy)+' • Maximum recommended '+money(maxBid)+' • Walk away above '+money(walkAway)+'.';}
      else if(current<=targetBuy){score='🟢 STRONG BUY';message='Current price is at or below BAM Target Buy. Estimated profit: '+money(Math.max(0,profit))+'.';}
      else if(current<=maxBid){score='🟢 GOOD BUY';message='Current price is within BAM recommended buying range. Do not chase beyond '+money(maxBid)+' unless you accept extra risk.';}
      else if(current<=walkAway){score='🟠 CAUTION';message='Above BAM recommended bid. '+money(walkAway)+' is the absolute walk-away ceiling based on Quick-Sale Value.';}
      else {score='🔴 DO NOT BID';message='STOP: current price is '+money(current-walkAway)+' above BAM Walk-Away Price.';}
    }
    document.getElementById('calc_quick').textContent=money(quick);
    document.getElementById('calc_costs').textContent=money(costs);
    document.getElementById('calc_target').textContent=money(targetBuy);
    document.getElementById('calc_max').textContent=money(maxBid);
    document.getElementById('calc_walkaway').textContent=money(walkAway);
    document.getElementById('calc_profit').textContent=quick>0 ? (profit<0?'-':'')+money(Math.abs(profit)) : '$0';
    document.getElementById('calc_score').textContent=score;
    document.getElementById('calc_message').textContent=message;
    const maxInput=byName('max_bid'); if(maxInput && quick>0) maxInput.value=maxBid.toFixed(2);
  }
  if(calcBtn) calcBtn.addEventListener('click',calculate);
  if(marketBtn) marketBtn.addEventListener('click',()=>{if(quickEl && marketMid>0){quickEl.value=marketMid;calculate();}});
  ['quick_sale_value','auction_fees','transport_cost','repair_allowance','rego_ppsr_cost','other_costs','target_profit','boat_engine_cost','boat_hull_cost','boat_trailer_cost','current_bid','asking_price','negotiated_price'].forEach(name=>{const el=byName(name);if(el)el.addEventListener('input',calculate);});
  if(quickEl && num('quick_sale_value')>0) calculate();
})();

refreshAssetFields();
refreshVariants();
</script></body></html>
"""


@app.get("/auction-watch")
@login_required
def auction_watch():
    conn=db(); q=(request.args.get("q") or "").strip(); asset_type=(request.args.get("asset_type") or "").strip(); status=(request.args.get("status") or "").strip(); source=(request.args.get("source") or "").strip(); make=(request.args.get("make") or "").strip(); model=(request.args.get("model") or "").strip()
    sql="SELECT a.*,(SELECT p.filename FROM auction_photos p WHERE p.auction_vehicle_id=a.id ORDER BY p.id LIMIT 1) AS thumbnail FROM auction_vehicles a WHERE 1=1"; params=[]
    if q:
        like=f"%{q}%"; sql+=" AND (make LIKE ? OR model LIKE ? OR variant LIKE ? OR auction_name LIKE ? OR auction_location LIKE ? OR lot_number LIKE ? OR listing_source LIKE ? OR seller_name LIKE ? OR seller_location LIKE ?)"; params += [like]*9
    if asset_type: sql+=" AND asset_type=?"; params.append(asset_type)
    if status: sql+=" AND status=?"; params.append(status)
    if source: sql+=" AND listing_source=?"; params.append(source)
    if make: sql+=" AND LOWER(make)=LOWER(?)"; params.append(make)
    if model: sql+=" AND LOWER(model)=LOWER(?)"; params.append(model)
    sql+=" ORDER BY CASE WHEN status IN ('Watching','Bidding') THEN 0 ELSE 1 END,COALESCE(auction_finish,'9999-12-31T23:59'),id DESC"
    rows=conn.execute(sql,params).fetchall(); conn.close()
    return render_template_string(AUCTION_PAGE,rows=rows,q=q,asset_type=asset_type,status=status,source_filter=source,make=make,model=model,types=AUCTION_TYPES,statuses=AUCTION_STATUSES,sources=BUYING_SOURCES)



@app.post("/auction-watch/import-listing")
@login_required
def auction_import_listing():
    data = request.get_json(silent=True) or {}
    url = (data.get("url") or "").strip()
    source_url = (data.get("source_url") or "").strip()
    source_hint = (data.get("source_hint") or "").strip()
    pasted = (data.get("text") or "").strip()
    try:
        if url:
            details = _fetch_listing_page(url)

        elif pasted:
            # The fallback box is for copied advertisement text. If a second URL
            # is pasted there, do not hit the blocked site again; explain what BAM needs.
            if pasted.lower().startswith(("http://", "https://")):
                if source_url:
                    raise ValueError("The lower box needs the advertisement text, not another link. Keep the link in the top box and paste the ad title, price, kilometres, location and description below.")
                details = _fetch_listing_page(pasted)
            else:
                # Keep the original listing URL as a source hint for pasted
                # Facebook/Gumtree/Carsales text without trying to fetch it again.
                details = _extract_listing_details(pasted, url=source_url)
                if source_hint in BUYING_SOURCES and source_hint != "Auction":
                    details["listing_source"] = source_hint
                if source_url:
                    details["listing_url"] = source_url
                    if details.get("listing_source") == "Auction":
                        details["auction_url"] = source_url
                elif details.get("listing_source") == "Other":
                    details["listing_source"] = "Facebook Marketplace"
        else:
            return jsonify(ok=False, error="Paste a listing link or listing text first."), 400
        return jsonify(ok=True, details=details)
    except ValueError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    except Exception as exc:
        return jsonify(ok=False, error=f"Import error: {exc}"), 400


@app.post("/auction-watch/automatic-valuation-preview")
@login_required
def auction_automatic_valuation_preview():
    try:
        details=request.get_json(silent=True) or {}
        result=bam_live_market_valuation(details)
        return jsonify(ok=True, **result)
    except Exception as exc:
        return jsonify(ok=False,error=str(exc)),400

@app.route("/auction-watch/add",methods=["GET","POST"])
@login_required
def auction_add():
    if request.method=="POST":
        conn=db()
        try:
            d=_auction_payload(); cols=list(d); cur=conn.execute(f"INSERT INTO auction_vehicles({','.join(cols)},created_by) VALUES({','.join(['?']*len(cols))},?)",list(d.values())+[session.get("display_name")]); aid=cur.lastrowid
            added=0
            for f in request.files.getlist("photos")[:10]:
                if f and f.filename:
                    fn=save_upload(f)
                    if fn:
                        conn.execute("INSERT INTO auction_photos(auction_vehicle_id,filename) VALUES(?,?)",(aid,fn))
                        added += 1
            if added < 10:
                _save_imported_listing_photos(conn, aid, request.form.get("imported_photo_urls"), referer=d.get("listing_url") or d.get("auction_url") or "", limit=10-added)
            conn.commit(); log_action("Auction vehicle added","auction_vehicle",aid,f"{d['make']} {d['model']}"); flash("Buying Watch vehicle saved.","success"); return redirect(url_for("auction_detail",auction_id=aid))
        except (ValueError,sqlite3.Error) as exc: conn.rollback(); flash(str(exc),"error")
        finally: conn.close()
    return render_template_string(AUCTION_FORM,item=None,photos=[],types=AUCTION_TYPES,statuses=AUCTION_STATUSES,sources=BUYING_SOURCES,conditions=AUCTION_CONDITIONS,today=date.today().isoformat(),makes=VEHICLE_MAKES,fuels=FUEL_TYPES,transmissions=TRANSMISSION_TYPES,drives=DRIVE_TYPES,model_catalog_json=json.dumps(VEHICLE_MODEL_CATALOG),variant_catalog_json=json.dumps(VEHICLE_VARIANT_CATALOG),asset_model_catalog_json=json.dumps(AUCTION_ASSET_MODEL_CATALOG))


@app.route("/auction-watch/<int:auction_id>",methods=["GET","POST"])
@login_required
def auction_detail(auction_id):
    conn=db(); item=conn.execute("SELECT * FROM auction_vehicles WHERE id=?",(auction_id,)).fetchone()
    if not item: conn.close(); return "Auction vehicle not found",404
    if request.method=="POST":
        try:
            d=_auction_payload(); conn.execute("UPDATE auction_vehicles SET "+",".join(f"{k}=?" for k in d)+",updated_at=? WHERE id=?",list(d.values())+[datetime.now().isoformat(timespec="seconds"),auction_id]); count=conn.execute("SELECT COUNT(*) c FROM auction_photos WHERE auction_vehicle_id=?",(auction_id,)).fetchone()["c"]
            added=0
            for f in request.files.getlist("photos")[:max(0,10-int(count))]:
                if f and f.filename:
                    fn=save_upload(f)
                    if fn:
                        conn.execute("INSERT INTO auction_photos(auction_vehicle_id,filename) VALUES(?,?)",(auction_id,fn))
                        added += 1
                        # If this auction vehicle has already been transferred to stock,
                        # keep newly-added Auction Watch photos in sync with that vehicle.
                        if item["won_vehicle_id"]:
                            ph=conn.execute("INSERT INTO vehicle_photos(vehicle_id,filename,caption) VALUES(?,?,?)",(item["won_vehicle_id"],fn,"Synced from Auction Watch"))
                            vehicle=conn.execute("SELECT featured_photo_id FROM vehicles WHERE id=?",(item["won_vehicle_id"],)).fetchone()
                            if vehicle and not vehicle["featured_photo_id"]:
                                conn.execute("UPDATE vehicles SET featured_photo_id=? WHERE id=?",(ph.lastrowid,item["won_vehicle_id"]))
            remaining=max(0,10-int(count)-added)
            if remaining:
                _save_imported_listing_photos(conn, auction_id, request.form.get("imported_photo_urls"), referer=d.get("listing_url") or d.get("auction_url") or "", limit=remaining)
            conn.commit(); log_action("Auction vehicle updated","auction_vehicle",auction_id,f"{d['make']} {d['model']}"); flash("Buying Watch vehicle updated.","success"); return redirect(url_for("auction_detail",auction_id=auction_id))
        except (ValueError,sqlite3.Error) as exc: conn.rollback(); flash(str(exc),"error")
    item=conn.execute("SELECT * FROM auction_vehicles WHERE id=?",(auction_id,)).fetchone(); photos=conn.execute("SELECT * FROM auction_photos WHERE auction_vehicle_id=? ORDER BY id",(auction_id,)).fetchall(); conn.close()
    market_query = " ".join(str(x) for x in (item["year"], item["make"], item["model"], item["variant"], item["fuel_type"], item["transmission"], item["drive_type"], (f"{item['odometer_km']}km" if item["odometer_km"] else "")) if x)
    carsales_valuation_url = "https://www.carsales.com.au/car-valuations/"
    carsales_search_url = "https://www.carsales.com.au/cars/?q=" + urllib.parse.quote(market_query)
    return render_template_string(AUCTION_FORM,item=item,photos=photos,types=AUCTION_TYPES,statuses=AUCTION_STATUSES,sources=BUYING_SOURCES,conditions=AUCTION_CONDITIONS,today=date.today().isoformat(),makes=VEHICLE_MAKES,fuels=FUEL_TYPES,transmissions=TRANSMISSION_TYPES,drives=DRIVE_TYPES,model_catalog_json=json.dumps(VEHICLE_MODEL_CATALOG),variant_catalog_json=json.dumps(VEHICLE_VARIANT_CATALOG),asset_model_catalog_json=json.dumps(AUCTION_ASSET_MODEL_CATALOG),carsales_valuation_url=carsales_valuation_url,carsales_search_url=carsales_search_url)


@app.get("/auction-watch/<int:auction_id>/market-value")
@login_required
def auction_value(auction_id):
    conn=db(); item=conn.execute("SELECT * FROM auction_vehicles WHERE id=?",(auction_id,)).fetchone()
    if not item: conn.close(); return "Auction vehicle not found",404
    low,mid,high,source,ac,rc=_auction_market_value(conn,item)
    dealer_low,dealer_high=low,high
    private_low,private_high=round(low*0.92,2),round(high*0.97,2)
    trade_low,trade_high=round(low*0.68,2),round(mid*0.78,2)
    wholesale_low,wholesale_high=round(low*0.62,2),round(mid*0.72,2)
    fees=float(item["auction_fees"] or 0)+float(item["transport_cost"] or 0)+float(item["other_costs"] or 0)
    suggested=max(0,round(wholesale_high-fees,2))
    confidence="Good" if (ac+rc)>=5 else "Limited" if (ac+rc)>0 else "Estimate only"
    provider="BAM internal history"
    detail=f"{source} ({ac} Buying Watch / {rc} retail comparables). Licensed live Australian provider not connected yet."
    conn.execute("""UPDATE auction_vehicles SET market_low=?,market_mid=?,market_high=?,private_value_low=?,private_value_high=?,wholesale_value_low=?,wholesale_value_high=?,trade_value_low=?,trade_value_high=?,dealer_value_low=?,dealer_value_high=?,suggested_buy_price=?,valuation_provider=?,valuation_confidence=?,valuation_source=?,valuation_checked_at=? WHERE id=?""",(low,mid,high,private_low,private_high,wholesale_low,wholesale_high,trade_low,trade_high,dealer_low,dealer_high,suggested,provider,confidence,detail,datetime.now().isoformat(timespec="seconds"),auction_id))
    conn.commit(); conn.close()
    flash("Valuation Hub updated. These figures use BAM history until your licensed live Australian valuation API is connected.","success")
    return redirect(url_for("auction_detail",auction_id=auction_id))


@app.post("/auction-watch/<int:auction_id>/automatic-valuation")
@login_required
def auction_automatic_valuation(auction_id):
    conn=db(); item=conn.execute("SELECT * FROM auction_vehicles WHERE id=?",(auction_id,)).fetchone()
    if not item: conn.close(); return "Auction vehicle not found",404
    try:
        result=bam_live_market_valuation(dict(item)); prices=(result["prices"]+[0,0,0,0,0])[:5]
        costs=sum(float(item[k] or 0) for k in ("auction_fees","transport_cost","repair_allowance","rego_ppsr_cost","other_costs"))
        if str(item["asset_type"] or "").lower()=="boat": costs += sum(float(item[k] or 0) for k in ("boat_engine_cost","boat_hull_cost","boat_trailer_cost"))
        suggested=max(0,result["wholesale_high"]-costs-float(item["target_profit"] or 0))
        source=result["summary"] + " Asking prices are market evidence, not confirmed sale prices."
        conn.execute("""UPDATE auction_vehicles SET comparable_price_1=?,comparable_price_2=?,comparable_price_3=?,comparable_price_4=?,comparable_price_5=?,market_low=?,market_mid=?,market_high=?,private_value_low=?,private_value_high=?,wholesale_value_low=?,wholesale_value_high=?,trade_value_low=?,trade_value_high=?,dealer_value_low=?,dealer_value_high=?,suggested_buy_price=?,quick_sale_value=?,valuation_provider=?,valuation_confidence=?,valuation_source=?,valuation_checked_at=? WHERE id=?""",(*prices,result["market_low"],result["market_mid"],result["market_high"],result["private_low"],result["private_high"],result["wholesale_low"],result["wholesale_high"],result["trade_low"],result["trade_high"],result["dealer_low"],result["dealer_high"],round(suggested,2),result["market_mid"],"BAM AI live Australian market search",result["confidence"],source,datetime.now().isoformat(timespec="seconds"),auction_id))
        conn.commit(); flash(f"BAM Automatic Market Valuation saved from {len(result['prices'])} current Australian comparable listing(s).","success")
    except Exception as exc:
        flash(f"Automatic market valuation could not complete: {exc}","error")
    finally: conn.close()
    return redirect(url_for("auction_detail",auction_id=auction_id))

@app.post("/auction-watch/<int:auction_id>/comparable-value")
@login_required
def auction_comparable_value(auction_id):
    conn = db()
    item = conn.execute("SELECT * FROM auction_vehicles WHERE id=?", (auction_id,)).fetchone()
    if not item:
        conn.close()
        return "Auction vehicle not found", 404
    prices = []
    stored = []
    for i in range(1, 6):
        raw = (request.form.get(f"comparable_price_{i}") or "").strip().replace(",", "").replace("$", "")
        try:
            value = max(0.0, float(raw)) if raw else 0.0
        except ValueError:
            value = 0.0
        stored.append(value)
        if value > 0:
            prices.append(value)
    if not prices:
        conn.close()
        flash("Enter at least one comparable advertised price first.", "error")
        return redirect(url_for("auction_detail", auction_id=auction_id))
    prices.sort()
    dealer_low, dealer_high = prices[0], prices[-1]
    market_mid = statistics.median(prices)
    private_low, private_high = dealer_low * .92, dealer_high * .97
    trade_low, trade_high = dealer_low * .68, market_mid * .78
    wholesale_low, wholesale_high = dealer_low * .62, market_mid * .72
    costs = sum(float(item[k] or 0) for k in ("auction_fees", "transport_cost", "repair_allowance", "rego_ppsr_cost", "other_costs"))
    if str(item["asset_type"] or "").lower() == "boat":
        costs += sum(float(item[k] or 0) for k in ("boat_engine_cost", "boat_hull_cost", "boat_trailer_cost"))
    suggested = max(0, wholesale_high - costs - float(item["target_profit"] or 0))
    # v25.16.1: automatically feed the comparable-market median into the Buy Calculator.
    quick_sale_value = round(market_mid, 2)
    confidence = "Good" if len(prices) >= 4 else "Limited" if len(prices) >= 2 else "Single comparable"
    detail = f"BAM estimate from {len(prices)} user-entered Australian advertised comparable price{'s' if len(prices) != 1 else ''}. Asking prices are not confirmed sale prices."
    conn.execute("""UPDATE auction_vehicles SET comparable_price_1=?,comparable_price_2=?,comparable_price_3=?,comparable_price_4=?,comparable_price_5=?,market_low=?,market_mid=?,market_high=?,private_value_low=?,private_value_high=?,wholesale_value_low=?,wholesale_value_high=?,trade_value_low=?,trade_value_high=?,dealer_value_low=?,dealer_value_high=?,suggested_buy_price=?,quick_sale_value=?,valuation_provider=?,valuation_confidence=?,valuation_source=?,valuation_checked_at=? WHERE id=?""",
        (*stored, round(dealer_low,2), round(market_mid,2), round(dealer_high,2), round(private_low,2), round(private_high,2), round(wholesale_low,2), round(wholesale_high,2), round(trade_low,2), round(trade_high,2), round(dealer_low,2), round(dealer_high,2), round(suggested,2), quick_sale_value, "BAM comparable market analysis", confidence, detail, datetime.now().isoformat(timespec="seconds"), auction_id))
    conn.commit()
    conn.close()
    flash(f"Market valuation calculated from {len(prices)} comparable price{'s' if len(prices) != 1 else ''}. Quick-Sale Value set automatically to ${quick_sale_value:,.0f}.", "success")
    return redirect(url_for("auction_detail", auction_id=auction_id))


@app.post("/auction-watch/<int:auction_id>/photos/<int:photo_id>/delete")
@login_required
def auction_delete_photo(auction_id,photo_id):
    conn=db(); conn.execute("DELETE FROM auction_photos WHERE id=? AND auction_vehicle_id=?",(photo_id,auction_id)); conn.commit(); conn.close(); flash("Auction photo removed.","success"); return redirect(url_for("auction_detail",auction_id=auction_id))


@app.post("/auction-watch/<int:auction_id>/delete")
@login_required
def auction_delete(auction_id):
    conn=db(); item=conn.execute("SELECT * FROM auction_vehicles WHERE id=?",(auction_id,)).fetchone()
    if not item: conn.close(); return "Auction vehicle not found",404
    conn.execute("DELETE FROM auction_vehicles WHERE id=?",(auction_id,)); conn.commit(); conn.close(); log_action("Auction vehicle deleted","auction_vehicle",auction_id,f"{item['make']} {item['model']}"); flash("Auction vehicle deleted.","success"); return redirect(url_for("auction_watch"))


@app.post("/auction-watch/<int:auction_id>/transfer-to-stock")
@login_required
def auction_transfer(auction_id):
    conn=db(); item=conn.execute("SELECT * FROM auction_vehicles WHERE id=?",(auction_id,)).fetchone()
    if not item: conn.close(); return "Auction vehicle not found",404
    if item["won_vehicle_id"]: conn.close(); flash("This auction vehicle is already in BAM Vehicle Stock.","error"); return redirect(url_for("auction_detail",auction_id=auction_id))
    try:
        ownership=(request.form.get("sale_ownership") or "BAM Joint").strip(); ownership=ownership if ownership in {"BAM Joint","Barry","Matt"} else "BAM Joint"; purchase_date=request.form.get("purchase_date") or date.today().isoformat(); price=float(item["sold_price"] or item["negotiated_price"] or item["current_bid"] or item["asking_price"] or item["max_bid"] or 0); landed=price+float(item["auction_fees"] or 0)+float(item["transport_cost"] or 0)+float(item["other_costs"] or 0); gst=round(landed/11,2) if landed else 0; stock=next_stock_number(conn); barry=landed if ownership=="Barry" else landed/2 if ownership=="BAM Joint" else 0; matt=landed if ownership=="Matt" else landed/2 if ownership=="BAM Joint" else 0; notes=f"Transferred from BAM Buying Watch. Source: {item['listing_source'] or 'Auction'}; Auction: {item['auction_name'] or '-'}; Lot: {item['lot_number'] or '-'}; Purchase price: ${price:,.2f}; Fees/transport/other included in landed cost."
        cur=conn.execute("""INSERT INTO vehicles(stock_no,status,purchase_date,make,model,variant,year,vin,registration,odometer_km,colour,purchase_price_inc_gst,purchase_gst,barry_contribution,matt_contribution,sale_ownership,notes,asset_type,engine_hours,fuel_type,drive_type,transmission_style) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(stock,"In Stock",purchase_date,item["make"],item["model"],item["variant"],item["year"],item["vin"],item["registration"],item["odometer_km"],item["colour"],landed,gst,barry,matt,ownership,notes,item["asset_type"],item["engine_hours"],item["fuel_type"],item["drive_type"],item["transmission"])); vid=cur.lastrowid; featured=None
        for p in conn.execute("SELECT * FROM auction_photos WHERE auction_vehicle_id=? ORDER BY id",(auction_id,)).fetchall():
            ph=conn.execute("INSERT INTO vehicle_photos(vehicle_id,filename,caption) VALUES(?,?,?)",(vid,p["filename"],"Transferred from Auction Watch")); featured=featured or ph.lastrowid
        if featured: conn.execute("UPDATE vehicles SET featured_photo_id=? WHERE id=?",(featured,vid))
        conn.execute("UPDATE auction_vehicles SET status='Won',won_vehicle_id=?,updated_at=? WHERE id=?",(vid,datetime.now().isoformat(timespec="seconds"),auction_id)); conn.commit(); log_action("Auction vehicle transferred to stock","vehicle",vid,f"Auction #{auction_id} -> {stock}"); flash(f"Vehicle added to BAM stock as {stock}. Auction details and photos were carried across.","success"); return redirect(url_for("vehicle_detail",vehicle_id=vid))
    except (ValueError,sqlite3.Error) as exc: conn.rollback(); flash(f"Could not transfer vehicle: {exc}","error"); return redirect(url_for("auction_detail",auction_id=auction_id))
    finally: conn.close()


@app.get("/health")
def health_check():
    return jsonify({"status": "ok", "app": APP_NAME, "version": APP_VERSION}), 200


@app.get("/ready")
def readiness_check():
    try:
        conn = db()
        conn.execute("SELECT 1").fetchone()
        conn.close()
        return jsonify({"status": "ready", "database": str(DB_PATH)}), 200
    except Exception as exc:
        app.logger.exception("Readiness check failed")
        return jsonify({"status": "not_ready", "error": str(exc)}), 503


# Gunicorn imports this module rather than executing it as __main__.
init_db()

if __name__ == "__main__":
    app.run(
        debug=os.environ.get("FLASK_DEBUG", "0") == "1",
        host=os.environ.get("BAM_HOST", "0.0.0.0"),
        port=int(os.environ.get("PORT", "5000")),
    )
