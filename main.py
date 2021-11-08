from flask import Flask,render_template,request
from sqlalchemy import create_engine
import pandas as pd
import os






df=pd.read_excel('data.xlsx')
# print(df.head())

app=Flask(__name__)
UPLOAD_FOLDER = 'document'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

engine=create_engine('sqlite:///data.db',echo=False)
df.to_sql('Database', con=engine, if_exists='append',index=False)

@app.route("/")
def main():
    return render_template("index.html")

@app.route("/", methods=["GET","POST"])
def excell_load():
    if request.method=="POST":
        document = request.files['document']
        excel =(document.excel)
        document.save(os.path.join(app.config['UPLOAD_FOLDER'], excel))

if __name__ == "__main__":
    app.run(debug=True)