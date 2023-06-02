from io import BytesIO
from flask import make_response
from flask import Flask, make_response, render_template, request, redirect, url_for, send_from_directory
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
# start app
app = Flask(__name__)
# config
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///Data.db'
# init db
db = SQLAlchemy(app)
# model


class Data(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50))
    email = db.Column(db.String(50))

    def __init__(self, name, email):
        self.name = name
        self.email = email


@app.route('/zahid')
def zahid():
    data = Data.query.all()
    return render_template('data.html', data=data)


@app.route('/')
def index():
    return render_template('upload.html')


@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']  # Get the uploaded file
    df = pd.read_excel(file)  # Read the Excel file into a pandas DataFrame

    # Create a list of Data objects to store in the database
    data_objects = []
    for index, row in df.iterrows():
        # Replace 'Name' with the actual column name from your Excel file
        name = row['name']
        # Replace 'Email' with the actual column name from your Excel file
        email = row['email']
        # Create a new Data object with the provided arguments
        data = Data(name, email)
        data_objects.append(data)

    # Store the data objects in the database
    db.session.bulk_save_objects(data_objects)
    db.session.commit()

    return 'Data uploaded and stored in the database!'


@app.route('/data')
def display_data():
    # Fetch the data from the database
    data = Data.query.all()

    # Create a pandas DataFrame from the fetched data
    df = pd.DataFrame([(entry.name, entry.email)
                      for entry in data], columns=['Name', 'Email'])

    # Convert DataFrame to Excel file format
    excel_file = BytesIO()
    df.to_excel(excel_file, index=False)
    excel_file.seek(0)

    # Create a response with the Excel file attachment
    response = make_response(excel_file.getvalue())
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    response.headers['Content-Disposition'] = 'attachment; filename=data.xlsx'

    return response


# run app
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
