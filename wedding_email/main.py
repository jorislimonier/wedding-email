import json
import os
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import jinja2
import numpy as np
import pandas as pd
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader
from unidecode import unidecode

from wedding_email.constants import *

load_dotenv()


def normalize_name(name: str) -> str:
    if pd.isna(name):
        return np.nan
    return unidecode(name.lower())


def load_guests(path_to_mariage_data: Path) -> pd.DataFrame:
    # Charger les invités
    guests = (
        pd.read_excel(
            path_to_mariage_data,
            engine="odf",
            sheet_name="Invités",
        )
        .iloc[:, :9]
        .dropna(how="all")
    )

    # Normaliser les noms
    guests.insert(
        loc=0,
        column="nom_complet",
        value=(guests["Prénom"] + " " + guests["Nom"]).apply(normalize_name),
    )

    return guests


def verify_adults_count(answers: pd.DataFrame) -> bool:
    nb_adults_computed = answers[
        [c for c in answers.columns if c.startswith("adulte_")]
    ].count(axis=1)
    return (nb_adults_computed == answers["nb_adultes"]).all()


def verify_children_count(answers: pd.DataFrame) -> bool:
    nb_children_computed = answers[
        [c for c in answers.columns if c.startswith("enfant_")]
    ].count(axis=1)
    return (nb_children_computed == answers["nb_enfants"]).all()


def load_answers(path_to_mariage_data: Path) -> pd.DataFrame:
    # Charger les réponses au questionnaire
    answers = (
        pd.read_excel(
            path_to_mariage_data,
            engine="odf",
            sheet_name="reponses_questionnaire",
        )
        .iloc[:, 2:14]
        .dropna(how="all")
    )

    cols_adults = [c for c in answers.columns if c.startswith("Adulte")]
    cols_children = [c for c in answers.columns if c.startswith("Enfant")]

    replace_cols = (
        {c: f"adulte_{i + 1}" for i, c in enumerate(cols_adults)}
        | {c: f"enfant_{i + 1}" for i, c in enumerate(cols_children)}
        | {
            "Combien d'adultes participeront au mariage ?  (>10 ans)": "nb_adultes",
            "Combien d'enfants participeront au mariage ? (entre 3 et 10 ans)": "nb_enfants",
        }
    )

    answers = answers.rename(columns=replace_cols)

    mail_dict = read_mail_dict()
    answers["courriel"] = answers["adulte_1"].map(mail_dict)

    assert verify_adults_count(answers), "Erreur dans le comptage des adultes"
    assert verify_children_count(answers), "Erreur dans le comptage des enfants"

    return answers


def read_mail_dict(
    path_to_mail_dict: Path = Path(RAW_DATA_PATH) / "adulte1_to_mail.json",
) -> dict:
    with open(path_to_mail_dict, "r") as f:
        mail_dict = json.load(f)
    return mail_dict


def populate_template(guest: pd.Series) -> None:
    env = Environment(loader=FileSystemLoader(INTERIM_DATA_PATH))
    template = env.get_template("template.html")

    adults = [
        guest[f"adulte_{idx}"]
        for idx in range(1, 6)
        if not pd.isna(guest[f"adulte_{idx}"])
    ]
    children = [
        guest[f"enfant_{idx}"]
        for idx in range(1, 6)
        if not pd.isna(guest[f"enfant_{idx}"])
    ]
    email_address = guest["courriel"]
    html = template.render(
        adults=adults, children=children, email_address=email_address
    )

    # Enregistrer le html généré
    processed_mail_path = PROCESSED_DATA_PATH / "mail.html"
    with open(processed_mail_path, "w") as f:
        f.write(html)


def generate_mail_text(guest: pd.Series) -> None:
    adults = [
        guest[f"adulte_{idx}"]
        for idx in range(1, 6)
        if not pd.isna(guest[f"adulte_{idx}"])
    ]
    children = [
        guest[f"enfant_{idx}"]
        for idx in range(1, 6)
        if not pd.isna(guest[f"enfant_{idx}"])
    ]
    email_address = guest["courriel"]

    text = f"""
    Bonjour,

    Le mariage n'est plus qu'à quelques semaines et nous sommes ravis de vous compter parmi nos invités.


    Nous devons confirmer au traiteur le nombre final de personnes qui seront présentes. Pourriez-vous confirmer qu'il n'y a pas de désistement et que les invités ci-dessous seront présents :

    Adultes : {", ".join(adults)}
    Enfants : {", ".join(children)}

    A bientôt,
    Hélène et Joris
    """

    processed_mail_path = PROCESSED_DATA_PATH / "mail.txt"
    with open(processed_mail_path, "w") as f:
        f.write(text)


def get_mail_html(path: Path = PROCESSED_DATA_PATH / "mail.html") -> str:
    with open(path, "r") as f:
        html = f.read()
    return html


def get_mail_text(path: Path = PROCESSED_DATA_PATH / "mail.txt") -> str:
    with open(path, "r") as f:
        text = f.read()
    return text


def send_email(
    sender_email: str,
    receiver_email: str,
    password: str,
    message: MIMEMultipart,
    text: str,
    html: str,
    cc: list[str] = ["joris.limonier@gmail.com", "helene.limonier@gmail.com"],
    bcc: list[str] = [],
) -> None:

    # Turn these into plain/html MIMEText objects
    part_text = MIMEText(_text=text, _subtype="plain")
    part_html = MIMEText(_text=html, _subtype="html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(payload=part_text)
    message.attach(payload=part_html)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(host="smtp.gmail.com", port=465, context=context) as server:
        server.login(user=sender_email, password=password)
        server.sendmail(
            from_addr=sender_email,
            to_addrs=[receiver_email] + cc + bcc,
            msg=message.as_string(),
        )


if __name__ == "__main__":

    sender_email = "jolimo1202@gmail.com"
    # receiver_email = "helene.limonier@gmail.com"
    receiver_email = "joris.limonier@hotmail.fr"

    message = MIMEMultipart("alternative")
    message["Subject"] = "Mariage Hélène et Joris | Liste de mariage et désistements"
    message["From"] = sender_email
    message["To"] = receiver_email

    text = get_mail_text()
    html = get_mail_html()

    send_email(
        sender_email=sender_email,
        receiver_email=receiver_email,
        password=MAIL_APP_PASSWORD,
        message=message,
        text=text,
        html=html,
    )
