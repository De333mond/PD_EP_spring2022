from flask import Flask, url_for, render_template

from main.main import getTable 
app = Flask(__name__)

OneZetHeight = 50;

@app.route('/')
def main():
    table = getTable()
    return render_template("base.html", table=table, zet=OneZetHeight)



if __name__ == "__main__":
    app.run()