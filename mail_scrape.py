import pandas as pd
from datetime import datetime
import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import gc



def get_body_msg(imap, messages, N):
    
    body_msg = []
    for i in range(messages, messages-N, -1):

        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)

                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain":
                            body_msg.append(body)
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        body_msg.append(body)
            gc.collect()
                        
    return body_msg


def clean_from_stars(my_str):
    my_str = my_str.strip("* ")
    my_str = my_str.replace(":*", ":")
    return my_str

def change_vaccin_1_second(my_list):
    temp_list = my_list.copy()
    for i in range(len(temp_list)-1, 0, -1):
        if "Nom vaccin 1" in temp_list[i]:
            temp_list[i] = temp_list[i].replace("Nom vaccin 1", "Nom vaccin 1_")
            break
    return temp_list

def preprocess_message(message):
    splitted_msg_body = message.split("\r\n")
    # remove empty strings from result:
    splitted_msg_body = [e for e in splitted_msg_body if e not in ['']]
    # let's merge the last element and before last --> those supposed to be only one element:
    new_splitted_msg_body = splitted_msg_body.copy()
    new_splitted_msg_body[-2] = new_splitted_msg_body[-2] + ' ' + new_splitted_msg_body[-1]
    del new_splitted_msg_body[-1]
    
    new_splitted_msg_body = [clean_from_stars(e) for e in new_splitted_msg_body]
    
    # Split items by "*":
    new_splitted_msg_body = [e.split("*") for e in new_splitted_msg_body]
    flatten_new_splitted_msg_body = [item for sublist in new_splitted_msg_body for item in sublist]
    flatten_new_splitted_msg_body = [e for e in flatten_new_splitted_msg_body if e not in ["", " "]]
    flatten_new_splitted_msg_body = [e.strip(" ") for e in flatten_new_splitted_msg_body]
    
    flatten_new_splitted_msg_body = change_vaccin_1_second(flatten_new_splitted_msg_body)

    gc.collect()
    
    return flatten_new_splitted_msg_body

def get_list_citoyen(list_body_msg):
    list_citoyen = []
    for temp_msg in list_body_msg:
        my_dict = {"Nom et Prenom citoyen":'', 
                   "Date de naissance citoyen":'', 
                   "Sexe":'', 
                   "Code inscription du citoyen":'', 
                   "Numéro de téléphone":'', 
                   "Addresse":'', 
                   "Type de réclamation":'', 
                   "Contenu":'',
                   "Date reclamation":'',
                   "Etat dossier":'',
                   "Atteint par le Covid-19":'',
                   "Nom vaccin 1":'',
                   "Nom vaccin 1_":'',
                   "Date vaccin 1":'',
                   "Date vaccin 2":'',
                   "Vaccin 1 Centre":'',
                   "Vaccin 2 Centre":'',
                  }
        for element in temp_msg:

            temp_element = element.split(":", 1)
            temp_key = temp_element[0].strip(" ")
            temp_value = temp_element[1].strip(" ")

            my_dict[temp_key] = temp_value



        list_citoyen.append(my_dict)
        gc.collect()
        
    return list_citoyen


def create_df(list_dict):
    return pd.DataFrame(list_dict)


def save_to_excel(df):
    path = "output_" + datetime.today().strftime("%Y_%m_%d") + ".xlsx"
    df.to_excel(path, index=False)
    
    return True






if __name__=="__main__":

    # account credentials
    username = "XXX@outlook.com"
    password = "XXX"

    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL("outlook.office365.com")
    # authenticate
    print("Attempting Login ...")
    auth_return = imap.login(username, password)
    print(auth_return)

    status, messages = imap.select("INBOX")
    # total number of emails
    messages = int(messages[0])
    print("We've found {} emails in your Inbox!".format(messages))

    # number of top emails to fetch --> FETCH ALL Emails:
    N = messages

    body_msg = get_body_msg(imap=imap, messages=messages, N=N)
    print("You have {} body messages to be processed ...".format(len(body_msg)))

    new_body_msg = [preprocess_message(e) for e in body_msg]

    list_citoyen = get_list_citoyen(list_body_msg=new_body_msg)
    print("The total Number of processed messages is : {}".format(len(list_citoyen)))

    citoyen_df = create_df(list_dict=list_citoyen)
    save_to_excel(df=citoyen_df)