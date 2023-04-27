# Importing required functions
import pandas as pd
from flask import Flask, render_template, request,redirect, url_for, send_from_directory
from fileinput import filename
import os
import warnings
import Calculadora
from Calculadora import calculadora

# Flask constructor
app = Flask(__name__)
 
# Root endpoint
@app.get('/')
def upload():
    return render_template('upload.html')
 
 
@app.post('/view')
def view():
 
    # Read the File using Flask request
    file = request.files['file']
    # save file in local directory
    file.save(file.filename)
   
 
    # Parse the data as a Pandas DataFrame type
    df=pd.read_excel(file, index_col=0)
    df.to_excel("DatosHotelEC.xlsx")
    #data = pandas.read_excel(file)
    import Calculadora
    output_file=calculadora(df)
    #output_file = "test2.docx"
    # Return HTML snippet that will render the table
    return output_file.to_html()
    #return render_template('Default.htm')

@app.route('/download')
def download():
    return render_template('download.html', files=os.listdir('images'))

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory('images', filename)
 
# Main Driver Function
if __name__ == '__main__':
    # Run the application on the local development server
    app.run(debug=True)

