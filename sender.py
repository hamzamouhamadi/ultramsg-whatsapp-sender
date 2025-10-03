import http.client
import ssl
import pandas as pd
from urllib.parse import urlencode
from dotenv import load_dotenv
import os

# Charger les variables d'environnement (.env)
load_dotenv()

# UltraMsg API credentials
TOKEN = os.getenv("ULTRAMSG_TOKEN")
INSTANCE_ID = os.getenv("ULTRAMSG_INSTANCE_ID")
API_URL = "api.ultramsg.com"

# Modèle de message si aucune colonne Message n’est fournie dans l’Excel.
MESSAGE_TEMPLATE = (
    "Bonjour {prenom},\n\n"
    "Je me permets de vous contacter dans le cadre du suivi des lauréats du programme JobInTech.\n"
    "Nous aimerions savoir où vous en êtes actuellement sur le plan professionnel, "
    "afin d’évaluer l’impact du programme et d’améliorer notre accompagnement.\n\n"
    "👉 Il vous suffit simplement de répondre par un court message :\n\n"
    "- Inséré : Oui / Non\n"
    "- Poste actuel (si Oui)\n\n"
    "Par ailleurs, sachez que nous pouvons également vous accompagner dans votre recherche d’opportunités "
    "en vous mettant en relation avec des entreprises partenaires.\n\n"
    "Votre retour est donc très important pour nous 🙏\n\n"
    "Bien cordialement,\n"
    "Chaimae Jerrar\n"
    "Cheffe de projet – JobInTech"
)


# --- LECTURE DU FICHIER EXCEL ---
def read_contacts(file_path, phone_col="PhoneNumber", prenom_col="Prenom", msg_col="Message"):
    try:
        df = pd.read_excel(file_path)

        if phone_col not in df.columns:
            raise ValueError(f"Colonne '{phone_col}' introuvable dans le fichier Excel.")

        # Ajouter colonnes vides si absentes
        if prenom_col not in df.columns:
            df[prenom_col] = None
        if msg_col not in df.columns:
            df[msg_col] = None

        return df[[phone_col, prenom_col, msg_col]].dropna(subset=[phone_col]).to_dict(orient="records")
    except Exception as e:
        print(f"Erreur de lecture du fichier Excel: {e}")
        return []

# --- ENVOI DU MESSAGE ---
def send_whatsapp_message(phone_number, prenom=None, message=None):
    try:
        conn = http.client.HTTPSConnection(API_URL, context=ssl._create_unverified_context())

        # Construire message
        if message and isinstance(message, str) and message.strip():
            message_body = message.strip()
        else:
            # utiliser le template avec remplacement des variables
            message_body = MESSAGE_TEMPLATE.format(
                prenom=prenom if prenom else "",
                phone=phone_number
            )

        payload = {
            "token": TOKEN,
            "to": phone_number,
            "body": message_body
        }
        payload_str = urlencode(payload)
        headers = {"Content-Type": "application/x-www-form-urlencoded"}

        endpoint = f"/{INSTANCE_ID}/messages/chat"
        conn.request("POST", endpoint, payload_str, headers)

        res = conn.getresponse()
        data = res.read().decode("utf-8")

        if res.status == 200:
            print(f"Message envoyé à {phone_number} ({prenom}): {data}")
            return True
        else:
            print(f"Échec d'envoi à {phone_number} ({prenom}): {data}")
            return False
    except Exception as e:
        print(f"Erreur envoi message à {phone_number} ({prenom}): {e}")
        return False
    finally:
        conn.close()

# --- MAIN ---
def main():
    excel_file = os.path.join(os.path.dirname(__file__), "contacts.xlsx")

    contacts = read_contacts(excel_file)

    if not contacts:
        print("Aucun contact trouvé ou erreur de lecture du fichier.")
        return

    success_count = 0
    for contact in contacts:
        phone = contact["PhoneNumber"]
        prenom = contact.get("Prenom")
        message = contact.get("Message")
        if send_whatsapp_message(phone, prenom, message):
            success_count += 1

    print(f"\nRésumé : {success_count}/{len(contacts)} messages envoyés avec succès.")

if __name__ == "__main__":
    main()
