import os
import pandas as pd
from flask import Flask, request, jsonify
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def process_excel(filepath):
    # Use pandas para ler o arquivo Excel
    df = pd.read_excel(filepath, engine='openpyxl')
    # Converta os dados para JSON
    excel_data = df.to_dict(orient='records')
    return excel_data


@app.route('/upload', methods=['POST'])
def upload_file():
    # Verifica se a solicitação tem o arquivo
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400

    file = request.files['file']

    # Se o usuário não selecionar nenhum arquivo, o navegador
    # enviará uma solicitação sem arquivo nomeado
    if file.filename == '':
        return jsonify({'error': 'Nenhum arquivo selecionado'}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # Processar o arquivo Excel
        excel_data = process_excel(filepath)

        # Criar listas dinamicamente para os dados desconhecidos
        dynamic_lists = {}
        for row in excel_data:
            for key, value in row.items():
                if pd.notna(value):
                    if key not in dynamic_lists:
                        dynamic_lists[key] = []
                    dynamic_lists[key].append(value)

        return jsonify({
            'message': 'Arquivo enviado com sucesso',
            'filename': filename,
            'dynamic_lists': dynamic_lists
        }), 200

    return jsonify({'error': 'Arquivo não suportado'}), 400


if __name__ == '__main__':
    app.run(debug=True)
