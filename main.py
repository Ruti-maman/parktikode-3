from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import base64
import os
import tempfile
import pythoncom
import win32com.client

# יצירת אפליקציית Flask
app = Flask(__name__)
CORS(app)  # מאפשר לדפדפן לשלוח בקשות ל-local server

# טעינת HTML דרך Flask
@app.route('/')
def index():
    return render_template('index.html')  # שמרי את HTML בתיקיית templates

# יצירת טיוטות Outlook
@app.route('/drafts', methods=['POST'])
def create_drafts():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No JSON data'}), 400

        subject = data.get('subject', '')
        body = data.get('body', '')
        recipients = data.get('recipients', [])
        file_obj = data.get('file', None)

        if not recipients:
            return jsonify({'error': 'No recipients provided'}), 400

        # טיפול בקובץ מצורף אם קיים
        attachment_path = None
        if file_obj and 'name' in file_obj and 'dataUrl' in file_obj:
            try:
                name = file_obj['name']
                data_url = file_obj['dataUrl']
                if ',' in data_url:
                    data_url = data_url.split(',')[1]
                attachment_bytes = base64.b64decode(data_url)
                fd, attachment_path = tempfile.mkstemp(suffix='_'+name)
                with os.fdopen(fd, 'wb') as f:
                    f.write(attachment_bytes)
            except Exception as e:
                return jsonify({'error': f'Failed to decode attachment: {str(e)}'}), 400

        # אתחול COM לפני שימוש ב-Outlook
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            return jsonify({'error': f'Cannot start Outlook: {str(e)}'}), 500

        # יצירת טיוטות לכל נמענים
        for recipient in recipients:
            try:
                mail = outlook.CreateItem(0)  # 0 = olMailItem
                mail.Subject = subject
                mail.Body = body
                mail.To = recipient
                if attachment_path:
                    mail.Attachments.Add(attachment_path)
                mail.Display()  # מציג את הטיוטה
            except Exception as e:
                return jsonify({'error': f'Failed to create draft for {recipient}: {str(e)}'}), 500

        # מחיקת הקובץ הזמני
        if attachment_path:
            os.remove(attachment_path)

        return jsonify({'status': 'ok', 'recipients_count': len(recipients)})

    except Exception as e:
        return jsonify({'error': f'Unexpected error: {str(e)}'}), 500

# הרצת Flask
if __name__ == '__main__':
    app.run(port=5000)
