import os
import sys
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
import pandas as pd
from flask import Flask, jsonify, request
from custom_logger import logger


app = Flask(__name__)

@app.route('/send/outlook/mail/', methods=['POST'])
def send_mail_from_outlook():
    try:
        pay_load = request.json
        logger.debug(pay_load)
    
        msg = MIMEMultipart()
        msg['From'] = pay_load['from']
        msg['To'] = ','.join(pay_load['to']) if type(pay_load['to']) == list else pay_load['to']
        if pay_load.get('cc'):
            msg['CC'] = ','.join(pay_load['cc']) if type(pay_load.get('cc')) == list else pay_load['cc']
        msg['Subject'] = pay_load['subject']
        
        with open(pay_load['body'], mode='r') as f2:
            Body = MIMEText(f2.read(), pay_load['body_type'] if pay_load.get('body_type') else 'plain', pay_load['encoding'] if pay_load.get('encoding') else 'utf-8')
        msg.attach(Body)
        
        if pay_load.get('custom_logo_path'):
            with open(pay_load['custom_logo_path'], mode='rb') as f1:
                image = MIMEImage(f1.read())
            image.add_header('Content-ID', '<logo>')
            msg.attach(image)
        
        if pay_load.get('attachments'):
            for file_path in list(pay_load['attachments']):
                with open(file_path, mode='rb') as f:
                    _attachment = MIMEApplication(f.read(), _subtype=file_path.split('.')[-1])
                _attachment.add_header('Content-Disposition', 'attachment', filename=os.path.split(file_path)[-1])
                msg.attach(_attachment)
        
        _context = ssl.create_default_context()
        with smtplib.SMTP(pay_load['host'] if pay_load.get('host') else 'smtp.office365.com' , pay_load['port'] if pay_load.get('port') else 587) as smtp:
            smtp.starttls(context=_context)
            smtp.login(pay_load['user_name'], pay_load['password'])
            smtp.send_message(msg)
            
    except Exception as ex:
        logger.error({'message': str((type(ex), sys.exc_info()[1], sys.exc_info()[2]))})
        return jsonify({'message': str((type(ex), sys.exc_info()[1], sys.exc_info()[2]))})
        
    else:
        logger.info({'message': 'mail sent successfully'})
        return jsonify({'message': 'mail sent successfully'})

        
if __name__ == '__main__':
    app.run()
    

