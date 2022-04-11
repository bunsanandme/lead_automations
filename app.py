import smtp
import imap
import sql
from flask import Flask, render_template
from config import Config
from flask_forms import ClientForm

app = Flask(__name__)
app.config.from_object(Config)


@app.route("/")
def welcome():
    return render_template("index.html")


@app.route("/get_client_data", methods=['GET', 'POST'])
def get_client_data():
    form = ClientForm()
    if form.validate_on_submit():
        output = sql.get_client_data(form.email.data)
        if output is None:
            return render_template("client_data.html", form=form, output="Такого клиента у нас нет!")
        output = " ".join(output)
        return render_template("client_data.html", form=form, output=output)
    return render_template("client_data.html", form=form, output="")
