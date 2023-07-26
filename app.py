import os
from pydub import AudioSegment
from pydub.playback import play
import smtplib
import email
import imaplib
import speech_recognition as sr
from gtts import gTTS
from email.header import decode_header
import logging
from playsound import playsound
from rapidfuzz import fuzz
from threading import Thread
 

# import webbrowser
from flask import Flask, render_template, request

app = Flask(__name__)

@app.route("/")
def homepage():
    return render_template("index.html")

@app.route("/composeMail")
def compose_mail():
    return render_template("composeMail.html")

@app.route("/Mail")
def search_mail():
    return render_template("Mail.html")

@app.route("/MailBoxStatus")
def get_mail_status():
    return render_template("MailBoxStatus.html")

@app.route("/mailinBin")
def get_latest_mail():
    return render_template("mailinBin.html")

@app.route("/mailinDrafts")
def mail_drafts():
    return render_template("mailinDrafts.html")

@app.route("/mailinImportantMails")
def mail_important():
    return render_template("mailinImportantMails.html")

@app.route("/mailinInbox")
def mail_inbox():
    return render_template("mailinInbox.html")

@app.route("/mailinSentMailbox")
def mail_sent():
    return render_template("mailinSentMailbox.html")

@app.route("/mailinSpam")
def spam_mail():
    return render_template("mailinSpam.html")

@app.route("/mailinStarredMails")
def starred_mail():
    return render_template("mailinStarredMails.html")


if __name__ == '__main__':
    from waitress import serve
    logging.basicConfig()
    logger = logging.getLogger('waitress')
    logger.setLevel(logging.DEBUG)
    serve(app,host="localhost", port=5000)