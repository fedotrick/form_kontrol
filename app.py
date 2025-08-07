import os
import json
from datetime import datetime
from flask import Flask, render_template, request, jsonify, redirect, url_for
import gspread
from google.oauth2.service_account import Credentials
from google.auth.exceptions import GoogleAuthError
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'your-secret-key-here')

# Google Sheets configuration
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

class GoogleSheetsManager:
    def __init__(self):
        self.client = None
        self.plavka_sheet = None
        self.control_sheet = None
        self.setup_credentials()
    
    def setup_credentials(self):
        """Setup Google Sheets credentials"""
        try:
            # Try to use service account from JSON file
            service_account_path = os.getenv('GOOGLE_SERVICE_ACCOUNT_PATH', 'service_account.json')
            if os.path.exists(service_account_path):
                creds = Credentials.from_service_account_file(service_account_path, scopes=SCOPES)
            else:
                # Try to use credentials from environment variable
                service_account_info = json.loads(os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON', '{}'))
                if service_account_info:
                    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
                else:
                    raise ValueError("No Google credentials found")
            
            self.client = gspread.authorize(creds)
            
            # Setup sheet IDs from environment variables
            plavka_sheet_id = os.getenv('PLAVKA_SHEET_ID')
            control_sheet_id = os.getenv('CONTROL_SHEET_ID')
            
            if plavka_sheet_id:
                self.plavka_sheet = self.client.open_by_key(plavka_sheet_id).sheet1
            
            if control_sheet_id:
                self.control_sheet = self.client.open_by_key(control_sheet_id).sheet1
            else:
                # Create or use existing control sheet
                self.control_sheet = self.create_control_sheet()
                
        except Exception as e:
            print(f"Error setting up Google Sheets: {e}")
            raise
    
    def create_control_sheet(self):
        """Create control sheet with proper headers"""
        try:
            # Try to find existing sheet or create new one
            spreadsheet_name = os.getenv('CONTROL_SHEET_NAME', 'Control Data')
            
            try:
                spreadsheet = self.client.open(spreadsheet_name)
            except gspread.SpreadsheetNotFound:
                spreadsheet = self.client.create(spreadsheet_name)
                # Share with current user
                spreadsheet.share('', perm_type='anyone', role='writer')
            
            sheet = spreadsheet.sheet1
            sheet.update_title('Control')
            
            # Setup headers
            headers = [
                'Номер_плавки', 'Контроль_отлито', 'Контроль_принято',
                'Контроль_дата_приемки', 'Контролер1', 'Контролер2',
                'Второй_сорт_раковины', 'Второй_сорт_зарез',
                'Доработка_раковины', 'Доработка_зарез',
                'Доработка_несоответствие_размеров', 'Доработка_несоответствие_внешнего_вида',
                'Доработка_наплыв_металла', 'Доработка_прорыв_металла',
                'Доработка_вырыв', 'Доработка_облой',
                'Доработка_песок_на_поверхности', 'Доработка_песок_в_резьбе',
                'Доработка_клей', 'Доработка_коробление',
                'Доработка_дефект_пеномодели', 'Доработка_лапы',
                'Доработка_питатель', 'Доработка_корона',
                'Доработка_смещение',
                'Окончательный_брак_недолив', 'Окончательный_брак_раковины',
                'Окончательный_брак_коробление', 'Окончательный_брак_спай',
                'Окончательный_брак_трещины', 'Окончательный_брак_пригар_песка',
                'Окончательный_брак_пористость', 'Окончательный_брак_вырыв',
                'Окончательный_брак_скол', 'Окончательный_брак_слом',
                'Окончательный_брак_зарез', 'Окончательный_брак_нарушение_геометрии',
                'Окончательный_брак_рыхлота', 'Окончательный_брак_непроклей',
                'Окончательный_брак_пеномодель', 'Окончательный_брак_наплыв_металла',
                'Окончательный_брак_несоответствие_размеров', 'Окончательный_брак_несоответствие_внешнего_вида',
                'Окончательный_брак_нарушение_маркировки', 'Окончательный_брак_неслитина',
                'Окончательный_брак_прочее',
                'Контролер3'
            ]
            
            if len(sheet.row_values(1)) < len(headers):
                sheet.insert_row(headers, 1)
            
            return sheet
            
        except Exception as e:
            print(f"Error creating control sheet: {e}")
            raise
    
    def get_available_plavka_numbers(self):
        """Get available plavka numbers from Google Sheets"""
        try:
            if not self.plavka_sheet:
                return []
            
            # Get all data from plavka sheet
            data = self.plavka_sheet.get_all_records()
            
            # Filter for '/25' numbers
            plavka_numbers = [row['Учетный_номер'] for row in data 
                           if '/25' in str(row.get('Учетный_номер', ''))]
            
            # Get already used numbers from control sheet
            if self.control_sheet:
                try:
                    control_data = self.control_sheet.get_all_records()
                    used_numbers = [str(row.get('Номер_плавки', '')) for row in control_data]
                    
                    # Filter out used numbers
                    available_numbers = [num for num in plavka_numbers 
                                       if str(num) not in used_numbers]
                except:
                    available_numbers = plavka_numbers
            else:
                available_numbers = plavka_numbers
            
            return sorted(available_numbers)
            
        except Exception as e:
            print(f"Error getting plavka numbers: {e}")
            return []
    
    def get_plavka_details(self, plavka_number):
        """Get details for a specific plavka number"""
        try:
            if not self.plavka_sheet:
                return None
            
            data = self.plavka_sheet.get_all_records()
            for row in data:
                if str(row.get('Учетный_номер', '')) == str(plavka_number):
                    return {
                        'Наименование_отливки': row.get('Наименование_отливки', '')
                    }
            
            return None
            
        except Exception as e:
            print(f"Error getting plavka details: {e}")
            return None
    
    def save_control_data(self, data):
        """Save control data to Google Sheets"""
        try:
            if not self.control_sheet:
                raise ValueError("Control sheet not available")
            
            # Prepare data row
            row_data = [
                data.get('номер_плавки', ''),
                data.get('контроль_отлито', ''),
                data.get('контроль_принято', ''),
                data.get('контроль_дата_приемки', ''),
                data.get('контролер1', ''),
                data.get('контролер2', ''),
                data.get('второй_сорт_раковины', ''),
                data.get('второй_сорт_зарез', ''),
                data.get('доработка_раковины', ''),
                data.get('доработка_зарез', ''),
                data.get('доработка_несоответствие_размеров', ''),
                data.get('доработка_несоответствие_внешнего_вида', ''),
                data.get('доработка_наплыв_металла', ''),
                data.get('доработка_прорыв_металла', ''),
                data.get('доработка_вырыв', ''),
                data.get('доработка_облой', ''),
                data.get('доработка_песок_на_поверхности', ''),
                data.get('доработка_песок_в_резьбе', ''),
                data.get('доработка_клей', ''),
                data.get('доработка_коробление', ''),
                data.get('доработка_дефект_пеномодели', ''),
                data.get('доработка_лапы', ''),
                data.get('доработка_питатель', ''),
                data.get('доработка_корона', ''),
                data.get('доработка_смещение', ''),
                data.get('окончательный_брак_недолив', ''),
                data.get('окончательный_брак_раковины', ''),
                data.get('окончательный_брак_коробление', ''),
                data.get('окончательный_брак_спай', ''),
                data.get('окончательный_брак_трещины', ''),
                data.get('окончательный_брак_пригар_песка', ''),
                data.get('окончательный_брак_пористость', ''),
                data.get('окончательный_брак_вырыв', ''),
                data.get('окончательный_брак_скол', ''),
                data.get('окончательный_брак_слом', ''),
                data.get('окончательный_брак_зарез', ''),
                data.get('окончательный_брак_нарушение_геометрии', ''),
                data.get('окончательный_брак_рыхлота', ''),
                data.get('окончательный_брак_непроклей', ''),
                data.get('окончательный_брак_пеномодель', ''),
                data.get('окончательный_брак_наплыв_металла', ''),
                data.get('окончательный_брак_несоответствие_размеров', ''),
                data.get('окончательный_брак_несоответствие_внешнего_вида', ''),
                data.get('окончательный_брак_нарушение_маркировки', ''),
                data.get('окончательный_брак_неслитина', ''),
                data.get('окончательный_брак_прочее', ''),
                data.get('контролер3', '')
            ]
            
            # Append the row
            self.control_sheet.append_row(row_data)
            return True
            
        except Exception as e:
            print(f"Error saving control data: {e}")
            raise

# Initialize Google Sheets manager
gs_manager = GoogleSheetsManager()

@app.route('/')
def index():
    """Main page"""
    return render_template('index.html')

@app.route('/api/plavka-numbers')
def get_plavka_numbers():
    """Get available plavka numbers"""
    try:
        numbers = gs_manager.get_available_plavka_numbers()
        return jsonify({'success': True, 'numbers': numbers})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/plavka-details/<plavka_number>')
def get_plavka_details(plavka_number):
    """Get details for a specific plavka number"""
    try:
        details = gs_manager.get_plavka_details(plavka_number)
        return jsonify({'success': True, 'details': details})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/save', methods=['POST'])
def save_data():
    """Save control data"""
    try:
        data = request.json
        gs_manager.save_control_data(data)
        return jsonify({'success': True, 'message': 'Данные успешно сохранены!'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/calculate', methods=['POST'])
def calculate_control_prinato():
    """Calculate control_prinato value"""
    try:
        data = request.json
        
        # Get all defect values
        defects = [
            int(data.get('второй_сорт_раковины', 0) or 0),
            int(data.get('второй_сорт_зарез', 0) or 0),
            int(data.get('доработка_раковины', 0) or 0),
            int(data.get('доработка_зарез', 0) or 0),
            int(data.get('доработка_несоответствие_размеров', 0) or 0),
            int(data.get('доработка_несоответствие_внешнего_вида', 0) or 0),
            int(data.get('доработка_наплыв_металла', 0) or 0),
            int(data.get('доработка_прорыв_металла', 0) or 0),
            int(data.get('доработка_вырыв', 0) or 0),
            int(data.get('доработка_облой', 0) or 0),
            int(data.get('доработка_песок_на_поверхности', 0) or 0),
            int(data.get('доработка_песок_в_резьбе', 0) or 0),
            int(data.get('доработка_клей', 0) or 0),
            int(data.get('доработка_коробление', 0) or 0),
            int(data.get('доработка_дефект_пеномодели', 0) or 0),
            int(data.get('доработка_лапы', 0) or 0),
            int(data.get('доработка_питатель', 0) or 0),
            int(data.get('доработка_корона', 0) or 0),
            int(data.get('доработка_смещение', 0) or 0),
            int(data.get('окончательный_брак_недолив', 0) or 0),
            int(data.get('окончательный_брак_раковины', 0) or 0),
            int(data.get('окончательный_брак_коробление', 0) or 0),
            int(data.get('окончательный_брак_спай', 0) or 0),
            int(data.get('окончательный_брак_трещины', 0) or 0),
            int(data.get('окончательный_брак_пригар_песка', 0) or 0),
            int(data.get('окончательный_брак_пористость', 0) or 0),
            int(data.get('окончательный_брак_вырыв', 0) or 0),
            int(data.get('окончательный_брак_скол', 0) or 0),
            int(data.get('окончательный_брак_слом', 0) or 0),
            int(data.get('окончательный_брак_зарез', 0) or 0),
            int(data.get('окончательный_брак_нарушение_геометрии', 0) or 0),
            int(data.get('окончательный_брак_рыхлота', 0) or 0),
            int(data.get('окончательный_брак_непроклей', 0) or 0),
            int(data.get('окончательный_брак_пеномодель', 0) or 0),
            int(data.get('окончательный_брак_наплыв_металла', 0) or 0),
            int(data.get('окончательный_брак_несоответствие_размеров', 0) or 0),
            int(data.get('окончательный_брак_несоответствие_внешнего_вида', 0) or 0),
            int(data.get('окончательный_брак_нарушение_маркировки', 0) or 0),
            int(data.get('окончательный_брак_неслитина', 0) or 0),
            int(data.get('окончательный_брак_прочее', 0) or 0)
        ]
        
        контроль_отлито = int(data.get('контроль_отлито', 0) or 0)
        контроль_принято = контроль_отлито - sum(defects)
        
        return jsonify({'success': True, 'result': max(0, контроль_принято)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)