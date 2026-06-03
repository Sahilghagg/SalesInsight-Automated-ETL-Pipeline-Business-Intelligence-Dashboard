"""
SalesInsight API with SQLite Database - Professional Version
Multi-page application with Login, Dashboard, Analytics, Sales Management,
Product Management, Customer Management, File Upload, and PDF Reports
"""

from flask import Flask, jsonify, request, render_template
from flask_cors import CORS
import os
import sqlite3
from datetime import datetime
from dotenv import load_dotenv
import pandas as pd
from werkzeug.utils import secure_filename
import tempfile

# Try to import weasyprint for PDF generation (optional)
# PDF generation disabled - use CSV export instead
WEASYPRINT_AVAILABLE = False

# Load environment variables
load_dotenv()

# Create Flask app
app = Flask(__name__)
CORS(app)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-secret-key')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# ==================== DATABASE SETUP ====================

DATABASE_PATH = 'salesinsight.db'
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def get_db():
    """Get database connection"""
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def init_db():
    """Initialize database tables"""
    conn = get_db()
    cursor = conn.cursor()
    
    # Create users table with role
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            first_name TEXT,
            last_name TEXT,
            role TEXT DEFAULT 'user',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Create sales table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            transaction_id TEXT UNIQUE NOT NULL,
            product_name TEXT NOT NULL,
            category TEXT,
            quantity REAL NOT NULL,
            price REAL NOT NULL,
            revenue REAL NOT NULL,
            region TEXT NOT NULL,
            date TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Create products table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_code TEXT UNIQUE NOT NULL,
            product_name TEXT NOT NULL,
            category TEXT,
            unit_price REAL NOT NULL,
            stock_quantity INTEGER DEFAULT 0,
            reorder_level INTEGER DEFAULT 10,
            is_active BOOLEAN DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Create customers table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_code TEXT UNIQUE NOT NULL,
            customer_name TEXT NOT NULL,
            email TEXT,
            phone TEXT,
            address TEXT,
            city TEXT,
            state TEXT,
            total_purchases REAL DEFAULT 0,
            last_purchase_date TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Insert demo user if no users exist
    cursor.execute('SELECT COUNT(*) as count FROM users')
    count = cursor.fetchone()['count']
    
    if count == 0:
        # Create demo admin user (password: admin123)
        cursor.execute('''
            INSERT INTO users (username, email, password, first_name, last_name, role)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', ('admin', 'admin@example.com', 'admin123', 'Admin', 'User', 'admin'))
        print("✅ Demo admin user created: admin / admin123")
        
        # Create demo regular user
        cursor.execute('''
            INSERT INTO users (username, email, password, first_name, last_name, role)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', ('user', 'user@example.com', 'user123', 'Regular', 'User', 'user'))
        print("✅ Demo regular user created: user / user123")
    
    # Insert sample sales if no sales exist
    cursor.execute('SELECT COUNT(*) as count FROM sales')
    sales_count = cursor.fetchone()['count']
    
    if sales_count == 0:
        sample_sales = [
            ('TXN-000001', 'Laptop', 'Electronics', 2, 50000, 100000, 'North', '2024-06-01'),
            ('TXN-000002', 'Mobile Phone', 'Electronics', 5, 25000, 125000, 'South', '2024-06-02'),
            ('TXN-000003', 'T-Shirt', 'Clothing', 10, 599, 5990, 'East', '2024-06-02'),
            ('TXN-000004', 'Office Chair', 'Furniture', 3, 5999, 17997, 'West', '2024-06-03'),
            ('TXN-000005', 'Headphones', 'Electronics', 8, 1999, 15992, 'North', '2024-06-03'),
            ('TXN-000006', 'Jeans', 'Clothing', 6, 1299, 7794, 'South', '2024-06-04'),
            ('TXN-000007', 'Tablet', 'Electronics', 4, 30000, 120000, 'Central', '2024-06-04'),
            ('TXN-000008', 'Shoes', 'Clothing', 3, 2499, 7497, 'West', '2024-06-05'),
        ]
        
        for sale in sample_sales:
            cursor.execute('''
                INSERT INTO sales (transaction_id, product_name, category, quantity, price, revenue, region, date)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', sale)
        print("✅ Sample sales data added!")
    
    # Insert sample products if no products exist
    cursor.execute('SELECT COUNT(*) as count FROM products')
    products_count = cursor.fetchone()['count']
    
    if products_count == 0:
        sample_products = [
            ('PRD-001', 'Laptop', 'Electronics', 50000, 25, 5, 1),
            ('PRD-002', 'Mobile Phone', 'Electronics', 25000, 50, 10, 1),
            ('PRD-003', 'T-Shirt', 'Clothing', 599, 200, 20, 1),
            ('PRD-004', 'Office Chair', 'Furniture', 5999, 15, 5, 1),
            ('PRD-005', 'Headphones', 'Electronics', 1999, 75, 15, 1),
            ('PRD-006', 'Jeans', 'Clothing', 1299, 100, 20, 1),
            ('PRD-007', 'Tablet', 'Electronics', 30000, 30, 8, 1),
            ('PRD-008', 'Shoes', 'Clothing', 2499, 80, 10, 1),
        ]
        
        for product in sample_products:
            cursor.execute('''
                INSERT INTO products (product_code, product_name, category, unit_price, stock_quantity, reorder_level, is_active)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', product)
        print("✅ Sample products added!")
    
    # Insert sample customers if no customers exist
    cursor.execute('SELECT COUNT(*) as count FROM customers')
    customers_count = cursor.fetchone()['count']
    
    if customers_count == 0:
        sample_customers = [
            ('CUST-001', 'Rajesh Kumar', 'rajesh@example.com', '9876543210', 'Delhi', 'Delhi', 'Delhi', 0, None),
            ('CUST-002', 'Priya Singh', 'priya@example.com', '9876543211', 'Mumbai', 'Mumbai', 'Maharashtra', 0, None),
            ('CUST-003', 'Amit Patel', 'amit@example.com', '9876543212', 'Bangalore', 'Bangalore', 'Karnataka', 0, None),
            ('CUST-004', 'Neha Sharma', 'neha@example.com', '9876543213', 'Chennai', 'Chennai', 'Tamil Nadu', 0, None),
            ('CUST-005', 'Vikram Reddy', 'vikram@example.com', '9876543214', 'Hyderabad', 'Hyderabad', 'Telangana', 0, None),
        ]
        
        for customer in sample_customers:
            cursor.execute('''
                INSERT INTO customers (customer_code, customer_name, email, phone, address, city, state, total_purchases, last_purchase_date)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', customer)
        print("✅ Sample customers added!")
    
    conn.commit()
    conn.close()
    print("✅ Database initialized successfully!")

# Initialize database when app starts
init_db()

# ==================== PAGE ROUTES ====================

@app.route('/')
def login_page():
    """Login/Register page"""
    return render_template('login.html')

@app.route('/dashboard')
def dashboard_page():
    """Main dashboard page"""
    return render_template('dashboard.html')

@app.route('/analytics')
def analytics_page():
    """Analytics page with charts"""
    return render_template('analytics.html')

@app.route('/sales')
def sales_page():
    """Sales management page"""
    return render_template('sales.html')

@app.route('/profile')
def profile_page():
    """User profile page"""
    return render_template('profile.html')

@app.route('/settings')
def settings_page():
    """Settings page"""
    return render_template('settings.html')

@app.route('/products')
def products_page():
    """Products management page"""
    return render_template('products.html')

@app.route('/customers')
def customers_page():
    """Customers management page"""
    return render_template('customers.html')

# ==================== HEALTH CHECK ====================

@app.route('/api/health', methods=['GET'])
def health_check():
    """Check if API is running"""
    return jsonify({
        'status': 'healthy',
        'message': 'SalesInsight API is running!',
        'timestamp': datetime.now().isoformat()
    }), 200

# ==================== USER ENDPOINTS ====================

@app.route('/api/users/register', methods=['POST'])
def register_user():
    """Register a new user"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        username = data.get('username')
        email = data.get('email')
        password = data.get('password')
        
        if not username or not email or not password:
            return jsonify({'error': 'Username, email, and password are required'}), 400
        
        conn = get_db()
        cursor = conn.cursor()
        
        # Check if user exists
        cursor.execute('SELECT id FROM users WHERE username = ? OR email = ?', (username, email))
        if cursor.fetchone():
            conn.close()
            return jsonify({'error': 'Username or email already exists'}), 409
        
        # Insert new user
        cursor.execute('''
            INSERT INTO users (username, email, password, first_name, last_name, role)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (username, email, password, data.get('first_name', ''), data.get('last_name', ''), 'user'))
        
        conn.commit()
        user_id = cursor.lastrowid
        conn.close()
        
        return jsonify({
            'message': 'User registered successfully',
            'user': {
                'id': user_id,
                'username': username,
                'email': email,
                'first_name': data.get('first_name', ''),
                'last_name': data.get('last_name', '')
            }
        }), 201
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/users', methods=['GET'])
def get_users():
    """Get all users"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT id, username, email, first_name, last_name, role, created_at FROM users')
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    return jsonify({
        'users': users,
        'count': len(users)
    }), 200

@app.route('/api/users/login', methods=['POST'])
def login_user():
    """Login endpoint"""
    try:
        data = request.get_json()
        username = data.get('username')
        password = data.get('password')
        
        if not username or not password:
            return jsonify({'error': 'Username and password required'}), 400
        
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM users WHERE username = ? AND password = ?', (username, password))
        user = cursor.fetchone()
        conn.close()
        
        if user:
            return jsonify({
                'message': 'Login successful',
                'user': dict(user)
            }), 200
        else:
            return jsonify({'error': 'Invalid credentials'}), 401
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/users/update', methods=['PUT'])
def update_profile():
    """Update user profile"""
    try:
        data = request.get_json()
        username = data.get('username')
        
        if not username:
            return jsonify({'error': 'Username required'}), 400
        
        conn = get_db()
        cursor = conn.cursor()
        
        updates = []
        values = []
        
        if data.get('first_name') is not None:
            updates.append('first_name = ?')
            values.append(data['first_name'])
        if data.get('last_name') is not None:
            updates.append('last_name = ?')
            values.append(data['last_name'])
        if data.get('email') is not None:
            updates.append('email = ?')
            values.append(data['email'])
        
        if updates:
            values.append(username)
            query = f"UPDATE users SET {', '.join(updates)} WHERE username = ?"
            cursor.execute(query, values)
            conn.commit()
        
        # Get updated user
        cursor.execute('SELECT id, username, email, first_name, last_name, role FROM users WHERE username = ?', (username,))
        user = cursor.fetchone()
        conn.close()
        
        return jsonify({
            'message': 'Profile updated successfully',
            'user': dict(user)
        }), 200
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/users/change-password', methods=['POST'])
def change_password():
    """Change user password"""
    try:
        data = request.get_json()
        username = data.get('username')
        current_password = data.get('current_password')
        new_password = data.get('new_password')
        
        conn = get_db()
        cursor = conn.cursor()
        
        # Verify current password
        cursor.execute('SELECT password FROM users WHERE username = ?', (username,))
        user = cursor.fetchone()
        
        if not user or user['password'] != current_password:
            conn.close()
            return jsonify({'error': 'Current password is incorrect'}), 401
        
        # Update password
        cursor.execute('UPDATE users SET password = ? WHERE username = ?', (new_password, username))
        conn.commit()
        conn.close()
        
        return jsonify({'message': 'Password changed successfully'}), 200
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/users/delete', methods=['DELETE'])
def delete_user_account():
    """Delete user account"""
    try:
        data = request.get_json()
        username = data.get('username')
        
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM users WHERE username = ?', (username,))
        conn.commit()
        conn.close()
        
        return jsonify({'message': 'Account deleted successfully'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ==================== ADMIN ENDPOINTS ====================

@app.route('/api/admin/users', methods=['GET'])
def get_all_users():
    """Get all users (admin only)"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT id, username, email, first_name, last_name, role, created_at FROM users')
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({'users': users}), 200

@app.route('/api/admin/users/<int:user_id>/role', methods=['PUT'])
def update_user_role(user_id):
    """Update user role (admin only)"""
    try:
        data = request.get_json()
        new_role = data.get('role')
        
        if new_role not in ['admin', 'user']:
            return jsonify({'error': 'Invalid role'}), 400
        
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('UPDATE users SET role = ? WHERE id = ?', (new_role, user_id))
        conn.commit()
        conn.close()
        
        return jsonify({'message': 'User role updated successfully'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/admin/clear-data', methods=['POST'])
def clear_all_data():
    """Clear all sales data"""
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM sales')
        # Reset auto increment
        cursor.execute("DELETE FROM sqlite_sequence WHERE name='sales'")
        conn.commit()
        conn.close()
        return jsonify({'message': 'All data cleared successfully'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ==================== SALES ENDPOINTS ====================

@app.route('/api/sales', methods=['POST'])
def add_sale():
    """Add a sales record"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        required_fields = ['product_name', 'quantity', 'price', 'region']
        for field in required_fields:
            if field not in data:
                return jsonify({'error': f'Missing required field: {field}'}), 400
        
        # Calculate revenue
        quantity = float(data['quantity'])
        price = float(data['price'])
        revenue = quantity * price
        
        # Get next ID for transaction
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) as count FROM sales')
        count = cursor.fetchone()['count']
        
        sale = {
            'transaction_id': f'TXN-{count + 1:06d}',
            'product_name': data['product_name'],
            'category': data.get('category', 'Uncategorized'),
            'quantity': quantity,
            'price': price,
            'revenue': revenue,
            'region': data['region'],
            'date': data.get('date', datetime.now().strftime('%Y-%m-%d'))
        }
        
        cursor.execute('''
            INSERT INTO sales (transaction_id, product_name, category, quantity, price, revenue, region, date)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (sale['transaction_id'], sale['product_name'], sale['category'], 
              sale['quantity'], sale['price'], sale['revenue'], sale['region'], sale['date']))
        
        conn.commit()
        sale_id = cursor.lastrowid
        conn.close()
        
        return jsonify({
            'message': 'Sale recorded successfully',
            'sale': {'id': sale_id, **sale}
        }), 201
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/sales', methods=['GET'])
def get_sales():
    """Get all sales records"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM sales ORDER BY id DESC')
    sales = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    return jsonify({
        'sales': sales,
        'count': len(sales)
    }), 200

@app.route('/api/sales/filter', methods=['POST'])
def filter_sales():
    """Get sales filtered by date range"""
    try:
        data = request.get_json()
        from_date = data.get('from_date')
        to_date = data.get('to_date')
        
        conn = get_db()
        cursor = conn.cursor()
        
        if from_date and to_date:
            cursor.execute('''
                SELECT * FROM sales 
                WHERE date BETWEEN ? AND ? 
                ORDER BY id DESC
            ''', (from_date, to_date))
        elif from_date:
            cursor.execute('''
                SELECT * FROM sales 
                WHERE date >= ? 
                ORDER BY id DESC
            ''', (from_date,))
        elif to_date:
            cursor.execute('''
                SELECT * FROM sales 
                WHERE date <= ? 
                ORDER BY id DESC
            ''', (to_date,))
        else:
            cursor.execute('SELECT * FROM sales ORDER BY id DESC')
        
        sales = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return jsonify({
            'sales': sales,
            'count': len(sales)
        }), 200
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/sales/upload', methods=['POST'])
def upload_sales_file():
    """Upload CSV/Excel file to bulk import sales"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'File type not allowed. Use CSV or Excel'}), 400
        
        # Save file
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        
        # Read file based on extension
        ext = filename.rsplit('.', 1)[1].lower()
        
        if ext == 'csv':
            df = pd.read_csv(filepath)
        else:
            df = pd.read_excel(filepath)
        
        # Process each row
        conn = get_db()
        cursor = conn.cursor()
        
        cursor.execute('SELECT COUNT(*) as count FROM sales')
        count = cursor.fetchone()['count']
        
        imported_count = 0
        errors = []
        
        for index, row in df.iterrows():
            try:
                # Map columns (support different column names)
                product_name = row.get('product_name') or row.get('Product Name') or row.get('product')
                quantity = float(row.get('quantity') or row.get('Quantity') or 0)
                price = float(row.get('price') or row.get('Price') or 0)
                region = row.get('region') or row.get('Region') or 'Unknown'
                category = row.get('category') or row.get('Category') or 'Uncategorized'
                sale_date = row.get('date') or row.get('Date') or datetime.now().strftime('%Y-%m-%d')
                
                if not product_name or quantity <= 0 or price <= 0:
                    errors.append(f"Row {index + 2}: Invalid data")
                    continue
                
                revenue = quantity * price
                count += 1
                transaction_id = f"TXN-{count:06d}"
                
                cursor.execute('''
                    INSERT INTO sales (transaction_id, product_name, category, quantity, price, revenue, region, date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (transaction_id, product_name, category, quantity, price, revenue, region, sale_date))
                
                imported_count += 1
                
            except Exception as e:
                errors.append(f"Row {index + 2}: {str(e)}")
        
        conn.commit()
        conn.close()
        
        # Clean up uploaded file
        os.remove(filepath)
        
        return jsonify({
            'message': f'Successfully imported {imported_count} records',
            'imported': imported_count,
            'errors': errors[:10]
        }), 200
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/sales/<int:sale_id>', methods=['DELETE'])
def delete_sale(sale_id):
    """Delete a sale by ID"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM sales WHERE id = ?', (sale_id,))
    conn.commit()
    affected = cursor.rowcount
    conn.close()
    
    if affected > 0:
        return jsonify({'message': 'Sale deleted successfully'}), 200
    return jsonify({'error': 'Sale not found'}), 404

# ==================== PRODUCT ENDPOINTS ====================

@app.route('/api/products', methods=['GET'])
def get_products():
    """Get all products"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM products WHERE is_active = 1 ORDER BY product_name')
    products = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({'products': products, 'count': len(products)}), 200

@app.route('/api/products/<int:product_id>', methods=['GET'])
def get_product(product_id):
    """Get single product"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM products WHERE id = ?', (product_id,))
    product = cursor.fetchone()
    conn.close()
    if product:
        return jsonify({'product': dict(product)}), 200
    return jsonify({'error': 'Product not found'}), 404

@app.route('/api/products', methods=['POST'])
def add_product():
    """Add new product"""
    try:
        data = request.get_json()
        
        required_fields = ['product_name', 'category', 'unit_price']
        for field in required_fields:
            if field not in data:
                return jsonify({'error': f'Missing field: {field}'}), 400
        
        conn = get_db()
        cursor = conn.cursor()
        
        # Generate product code
        cursor.execute('SELECT COUNT(*) as count FROM products')
        count = cursor.fetchone()['count']
        product_code = f"PRD-{count + 1:03d}"
        
        cursor.execute('''
            INSERT INTO products (product_code, product_name, category, unit_price, stock_quantity, reorder_level)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (product_code, data['product_name'], data['category'], 
              data['unit_price'], data.get('stock_quantity', 0), data.get('reorder_level', 10)))
        
        conn.commit()
        product_id = cursor.lastrowid
        conn.close()
        
        return jsonify({'message': 'Product added successfully', 'product_id': product_id}), 201
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/products/<int:product_id>', methods=['PUT'])
def update_product(product_id):
    """Update product"""
    try:
        data = request.get_json()
        conn = get_db()
        cursor = conn.cursor()
        
        updates = []
        values = []
        
        for field in ['product_name', 'category', 'unit_price', 'stock_quantity', 'reorder_level']:
            if field in data:
                updates.append(f"{field} = ?")
                values.append(data[field])
        
        if updates:
            values.append(product_id)
            cursor.execute(f"UPDATE products SET {', '.join(updates)} WHERE id = ?", values)
            conn.commit()
        
        conn.close()
        return jsonify({'message': 'Product updated successfully'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/products/<int:product_id>', methods=['DELETE'])
def delete_product(product_id):
    """Soft delete product"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('UPDATE products SET is_active = 0 WHERE id = ?', (product_id,))
    conn.commit()
    conn.close()
    return jsonify({'message': 'Product deleted successfully'}), 200

# ==================== CUSTOMER ENDPOINTS ====================

@app.route('/api/customers', methods=['GET'])
def get_customers():
    """Get all customers"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM customers ORDER BY customer_name')
    customers = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({'customers': customers, 'count': len(customers)}), 200

@app.route('/api/customers', methods=['POST'])
def add_customer():
    """Add new customer"""
    try:
        data = request.get_json()
        
        required_fields = ['customer_name']
        for field in required_fields:
            if field not in data:
                return jsonify({'error': f'Missing field: {field}'}), 400
        
        conn = get_db()
        cursor = conn.cursor()
        
        cursor.execute('SELECT COUNT(*) as count FROM customers')
        count = cursor.fetchone()['count']
        customer_code = f"CUST-{count + 1:03d}"
        
        cursor.execute('''
            INSERT INTO customers (customer_code, customer_name, email, phone, address, city, state)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (customer_code, data['customer_name'], data.get('email', ''), 
              data.get('phone', ''), data.get('address', ''), data.get('city', ''), data.get('state', '')))
        
        conn.commit()
        customer_id = cursor.lastrowid
        conn.close()
        
        return jsonify({'message': 'Customer added successfully', 'customer_id': customer_id}), 201
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/customers/<int:customer_id>', methods=['GET'])
def get_customer(customer_id):
    """Get single customer"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM customers WHERE id = ?', (customer_id,))
    customer = cursor.fetchone()
    conn.close()
    if customer:
        return jsonify({'customer': dict(customer)}), 200
    return jsonify({'error': 'Customer not found'}), 404

@app.route('/api/customers/<int:customer_id>', methods=['PUT'])
def update_customer(customer_id):
    """Update customer"""
    try:
        data = request.get_json()
        conn = get_db()
        cursor = conn.cursor()
        
        updates = []
        values = []
        
        for field in ['customer_name', 'email', 'phone', 'address', 'city', 'state']:
            if field in data:
                updates.append(f"{field} = ?")
                values.append(data[field])
        
        if updates:
            values.append(customer_id)
            cursor.execute(f"UPDATE customers SET {', '.join(updates)} WHERE id = ?", values)
            conn.commit()
        
        conn.close()
        return jsonify({'message': 'Customer updated successfully'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/customers/<int:customer_id>', methods=['DELETE'])
def delete_customer(customer_id):
    """Delete customer"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM customers WHERE id = ?', (customer_id,))
    conn.commit()
    conn.close()
    return jsonify({'message': 'Customer deleted successfully'}), 200

# ==================== PDF REPORT ENDPOINTS ====================
# ==================== CSV REPORT ENDPOINT ====================

@app.route('/api/reports/sales-csv', methods=['POST'])
def generate_sales_csv():
    """Generate CSV sales report (works without PDF library)"""
    try:
        data = request.get_json()
        from_date = data.get('from_date')
        to_date = data.get('to_date')
        
        conn = get_db()
        cursor = conn.cursor()
        
        # Get sales data
        if from_date and to_date:
            cursor.execute('''
                SELECT * FROM sales 
                WHERE date BETWEEN ? AND ? 
                ORDER BY date DESC
            ''', (from_date, to_date))
        else:
            cursor.execute('SELECT * FROM sales ORDER BY date DESC')
        
        sales = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        # Create CSV content
        import io
        output = io.StringIO()
        
        # Write header
        output.write("Date,Product,Category,Quantity,Price,Revenue,Region\n")
        
        # Write data
        for sale in sales:
            output.write(f"{sale['date']},{sale['product_name']},{sale['category']},{sale['quantity']},{sale['price']},{sale['revenue']},{sale['region']}\n")
        
        csv_content = output.getvalue()
        output.close()
        
        return jsonify({
            'message': 'CSV report generated successfully',
            'csv_content': csv_content,
            'record_count': len(sales)
        }), 200
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ==================== ANALYTICS ENDPOINTS ====================

@app.route('/api/sales/summary', methods=['GET'])
def sales_summary():
    """Get sales summary statistics"""
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT SUM(revenue) as total_revenue, COUNT(*) as total_orders, SUM(quantity) as total_products FROM sales')
    result = cursor.fetchone()
    conn.close()
    
    total_revenue = result['total_revenue'] or 0
    total_orders = result['total_orders'] or 0
    total_products = result['total_products'] or 0
    avg_order_value = total_revenue / total_orders if total_orders > 0 else 0
    
    return jsonify({
        'total_revenue': round(total_revenue, 2),
        'total_orders': total_orders,
        'average_order_value': round(avg_order_value, 2),
        'total_products_sold': total_products
    }), 200

@app.route('/api/sales/by-region', methods=['GET'])
def sales_by_region():
    """Get sales grouped by region"""
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT region, SUM(revenue) as revenue, COUNT(*) as orders, SUM(quantity) as quantity
        FROM sales
        GROUP BY region
        ORDER BY revenue DESC
    ''')
    
    results = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    return jsonify({
        'data': results,
        'total_regions': len(results)
    }), 200

@app.route('/api/sales/by-product', methods=['GET'])
def sales_by_product():
    """Get sales grouped by product"""
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT product_name, category, SUM(revenue) as revenue, SUM(quantity) as quantity, COUNT(*) as orders
        FROM sales
        GROUP BY product_name
        ORDER BY revenue DESC
        LIMIT 10
    ''')
    
    results = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    return jsonify({
        'data': results,
        'total_products': len(results)
    }), 200

@app.route('/api/sales/trend', methods=['GET'])
def sales_trend():
    """Get sales trend over time"""
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT date, SUM(revenue) as revenue, COUNT(*) as orders, SUM(quantity) as quantity
        FROM sales
        GROUP BY date
        ORDER BY date DESC
        LIMIT 7
    ''')
    
    results = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    return jsonify({
        'data': results[::-1],
        'period': 'last_7_days'
    }), 200

# ==================== ERROR HANDLERS ====================

@app.errorhandler(404)
def not_found(error):
    return jsonify({'error': 'Endpoint not found'}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({'error': 'Internal server error'}), 500

# ==================== RUN THE APP ====================

if __name__ == '__main__':
    print("\n" + "=" * 70)
    print("🚀 SALESINSIGHT API STARTING - PROFESSIONAL EDITION")
    print("=" * 70)
    print(f"📍 Login Page:        http://localhost:5000")
    print(f"📊 Dashboard:         http://localhost:5000/dashboard")
    print(f"📈 Analytics:         http://localhost:5000/analytics")
    print(f"💰 Sales:             http://localhost:5000/sales")
    print(f"📦 Products:          http://localhost:5000/products")
    print(f"👥 Customers:         http://localhost:5000/customers")
    print(f"👤 Profile:           http://localhost:5000/profile")
    print(f"⚙️ Settings:          http://localhost:5000/settings")
    print(f"🏥 Health Check:      http://localhost:5000/api/health")
    print(f"💾 Database:          {DATABASE_PATH}")
    print("=" * 70)
    print("\n✨ Demo Credentials:")
    print("   Admin:  admin / admin123")
    print("   User:   user / user123")
    print("\n📁 Upload folder:", UPLOAD_FOLDER)
    print("📄 PDF Support:", "✅ Enabled" if WEASYPRINT_AVAILABLE else "❌ Disabled (run: pip install weasyprint)")
    print("\n🚀 Ready to accept requests!\n")
    
    app.run(
        host='0.0.0.0',
        port=5000,
        debug=True
    )