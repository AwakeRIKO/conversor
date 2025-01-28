# Adicione estas importações no topo do arquivo
from flask import Flask, request, send_file, render_template, jsonify
import os
from werkzeug.utils import secure_filename

# Crie as pastas necessárias
if not os.path.exists('static'):
    os.makedirs('static')
if not os.path.exists('uploads'):
    os.makedirs('uploads')

# Configure o Flask para servir arquivos estáticos
app = Flask(__name__, static_folder='static')
