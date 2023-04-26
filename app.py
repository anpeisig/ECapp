# Importing required functions
import pandas as pd
from flask import Flask, render_template, request
from fileinput import filename
import os
import warnings


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
    
    # Return HTML snippet that will render the table
    return df.to_html()
    #return render_template('Default.htm')
 
 
# Main Driver Function
if __name__ == '__main__':
    # Run the application on the local development server
    app.run(debug=True)
