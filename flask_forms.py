from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, BooleanField, SubmitField, DateField
from wtforms.validators import DataRequired


class ClientForm(FlaskForm):
    email = StringField(validators=[DataRequired()])
    sub = SubmitField("Получить")

