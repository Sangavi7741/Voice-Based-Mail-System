#Imports
import os
import smtplib
import email
import imaplib
import speech_recognition as sr
from gtts import gTTS
from email.header import decode_header
import logging
from rapidfuzz import fuzz
import time
from playsound import playsound

# Selenium imports
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)


# Globals
EMAIL_ID = "demomailpriya@gmail.com"   # Put your email id within quotes
PASSWORD = "nkaideifczkcefqw"   # Put your password within quotes
LANGUAGE = "en"
localhost_url = "http://127.0.0.1:5000/"

class SpeechMail:
    def __init__(self) -> None:
        self.email = EMAIL_ID
        self.password = PASSWORD
        self.lang = LANGUAGE

    def SpeakText(self, command, langinp=LANGUAGE):
        """
        Text to Speech using GTTS

        Args:
            command (str): Text to speak
            langinp (str, optional): Output language. Defaults to "en".
        """
        if langinp == "":
            langinp = "en"
        tts = gTTS(text=command, lang=langinp)
        tts.save("tempfile01.mp3")
        file = "tempfile01.mp3"
        os.system("start /min " + file)
        print(command)
        #playsound(file)
        time.sleep(10)
        os.remove("tempfile01.mp3")

    def speech_to_text(self):
        """
        Speech to text

        Returns:
            str: Returns transcripted text
        """
        r = sr.Recognizer()
        try:
            with sr.Microphone() as source2:
                r.adjust_for_ambient_noise(source2, duration=0.2)
                print("speak")
                audio2 = r.listen(source2)
                MyText = r.recognize_google(audio2)
                print("You said: "+MyText)
                return MyText

        except sr.RequestError as e:
            print("Could not request results; {0}".format(e))
            return None

        except sr.UnknownValueError:
            print("unknown error occured")
            return None
    
    def sendMail(self, sendTo, subject,msg):
        """
        To send a mail

        Args:
            sendTo (list): List of mail targets
            msg (str): Message
        """
        mail = smtplib.SMTP('smtp.gmail.com', 587)  # host and port
        # Hostname to send for this command defaults to the FQDN of the local host.
        mail.ehlo()
        mail.starttls()  # security connection
        mail.login(EMAIL_ID, PASSWORD)  # login part
        for person in sendTo:
            email_msg = f"Subject: {subject}\n\n{msg}"
            mail.sendmail(EMAIL_ID, person, email_msg)  # send part
            print("Mail sent successfully to " + person)
        mail.close()
    
    def composeMail(self):
        """
        Compose and create a Mail

        Returns:
            None: None
        """
        self.SpeakText("Mention the gmail ID of the persons to whom you want to send a mail.")
        receivers = self.speech_to_text()
        receivers = receivers + '@gmail.com'
        emails = receivers.split(" and ")
        index = 0
        for email in emails:
            emails[index] = email.replace(" ", "")
            index += 1

        self.SpeakText("The mail will be send to " +
                (' and '.join([str(elem) for elem in emails])) + ". Confirm by saying YES or NO.")
        confirmMailList = self.speech_to_text()
        if fuzz.partial_ratio("yes", confirmMailList) >= 60 or fuzz.ratio("s", confirmMailList) >= 60:
            self.SpeakText("What is the subject of the email?")
            subject = self.speech_to_text()

            self.SpeakText("Say your message")
            msg = self.speech_to_text()

            self.SpeakText("You said  " + msg + ". Confirm by saying YES or NO.")
            confirmMailBody = self.speech_to_text()
            if fuzz.partial_ratio("yes", confirmMailBody) >= 60 or fuzz.ratio("s", confirmMailBody) >= 60:
                self.SpeakText("Message sent")
                self.sendMail(emails, subject, msg)
            else:
                self.SpeakText("Operation cancelled by the user")
                return None
        else:
                self.SpeakText("Operation cancelled by the user")
                return None
        """  
        if confirmMailList.lower() != "yes":
            self.SpeakText("Operation cancelled by the user")
            return None

        self.SpeakText("What is the subject of the email?")
        subject = self.speech_to_text()

        self.SpeakText("Say your message")
        msg = self.speech_to_text()

        self.SpeakText("You said  " + msg + ". Confirm by saying YES or NO.")
        confirmMailBody = self.speech_to_text()
        if confirmMailBody.lower() == "yes":
            self.SpeakText("Message sent")
            self.sendMail(emails, subject, msg)
        else:
            self.SpeakText("Operation cancelled by the user")
            return None
        """

    def getMailBoxStatus(self):
        """
        Get mail counts of all folders in the mailbox
        """
        # host and port (ssl security)
        M = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        M.login(EMAIL_ID, PASSWORD)  # login

        for i in M.list()[1]:
            l = i.decode().split(' "/" ')
            if l[1] == '"[Gmail]"':
                continue

            stat, total = M.select(f'{l[1]}')
            l[1] = l[1][1:-1]
            messages = int(total[0])
            if l[1] == 'INBOX':
                self.SpeakText(l[1] + " has " + str(messages) + " messages.")
            else:
                self.SpeakText(l[1].split("/")[-1] + " has " +
                        str(messages) + " messages.")

        M.close()
        M.logout()


    def clean(self, text):
        """
        clean text for creating a folder
        """
        return "".join(c if c.isalnum() else "_" for c in text)


    def getLatestMails(self):
        """
        Get latest mails from folders in mailbox (Defaults to 3 Inbox mails)
        """
        mailBoxTarget = "INBOX"
        self.SpeakText("To get the latest mail, Say 1 for Inbox. Say 2 for Sent Mailbox. Say 3 for Drafts. Say 4 for important mails. Say 5 for Spam. Say 6 for Starred Mails. Say 7 for Bin.")
        cmb = self.speech_to_text()
        cmb = cmb.split()[-1]
        if fuzz.partial_ratio("1", cmb) >= 60 or fuzz.ratio("one", cmb) >= 60:
            mailBoxTarget = "INBOX"
            driver.get(localhost_url + "mailinInbox")
            self.SpeakText("Inbox selected.")
        elif fuzz.partial_ratio("2", cmb) >= 60 or fuzz.ratio("two", cmb) >= 60 or cmb.lower() == "tu":
            mailBoxTarget = '"[Gmail]/Sent Mail"'
            driver.get(localhost_url + "mailinSentMailbox")
            self.SpeakText("Sent Mailbox selected.")
        elif fuzz.partial_ratio("3", cmb) >= 60 or fuzz.ratio("three", cmb) >= 60:
            mailBoxTarget = '"[Gmail]/Drafts"'
            driver.get(localhost_url + "mailinDrafts")
            self.SpeakText("Drafts selected.")
        elif fuzz.partial_ratio("4", cmb) >= 60 or fuzz.ratio("four", cmb) >= 60:
            mailBoxTarget = '"[Gmail]/Important"'
            driver.get(localhost_url + "mailinImportantMails")
            self.SpeakText("Important Mails selected.")
        elif fuzz.partial_ratio("5", cmb) >= 60 or fuzz.ratio("five", cmb) >= 60:
            mailBoxTarget = '"[Gmail]/Spam"'
            driver.get(localhost_url + "mailinSpam")
            self.SpeakText("Spam selected.")
        elif fuzz.partial_ratio("6", cmb) >= 60 or fuzz.ratio("six", cmb) >= 60:
            mailBoxTarget = '"[Gmail]/Starred"'
            driver.get(localhost_url + "mailinStarredMails")
            self.SpeakText("Starred Mails selected.")
        elif fuzz.partial_ratio("7", cmb) >= 60 or fuzz.ratio("seven", cmb) >= 60:
            mailBoxTarget = '"[Gmail]/Bin"'
            driver.get(localhost_url + "mailinBin")
            self.SpeakText("Bin selected.")
        else:
            driver.get(localhost_url + "mailinInbox")
            self.SpeakText("Wrong choice. Hence, default option Inbox wil be selected.")

        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(EMAIL_ID, PASSWORD)

        status, messages = imap.select(mailBoxTarget)

        messages = int(messages[0])

        if messages == 0:
            self.SpeakText("Selected MailBox is empty.")
            return None
        elif messages == 1:
            N = 1   # number of top emails to fetch
        elif messages == 2:
            N = 2   # number of top emails to fetch
        else:
            N = 3   # number of top emails to fetch

        msgCount = 1
        for i in range(messages, messages-N, -1):
            self.SpeakText(f"Message {msgCount}:")
            # fetch the email message by ID
            res, msg = imap.fetch(str(i), "(RFC822)")
            for response in msg:
                if isinstance(response, tuple):
                    # parse a bytes email into a message object
                    msg = email.message_from_bytes(response[1])

                    subject, encoding = decode_header(msg["Subject"])[
                        0]    # decode the email subject
                    if isinstance(subject, bytes):
                        # if it's a bytes, decode to str
                        subject = subject.decode(encoding)

                    From, encoding = decode_header(msg.get("From"))[
                        0]      # decode email sender
                    if isinstance(From, bytes):
                        From = From.decode(encoding)
                    self.SpeakText("Subject: " + subject)
                    FromArr = From.split()
                    FromName = " ".join(namechar for namechar in FromArr[0:-1])
                    self.SpeakText("From: " + FromName)
                    self.SpeakText("Sender mail: " + FromArr[-1])
                    self.SpeakText("The mail says or contains the following:")

                    # MULTIPART
                    if msg.is_multipart():
                        for part in msg.walk():  # iterate over email parts
                            content_type = part.get_content_type()      # extract content type of email
                            content_disposition = str(
                                part.get("Content-Disposition"))
                            try:
                                body = part.get_payload(
                                    decode=True).decode()   # get the email body
                            except:
                                pass

                            # PLAIN TEXT MAIL
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                self.SpeakText(
                                    "Do you want to listen to the text content of the mail ? Please say YES or NO.")
                                talkMSG1 = self.speech_to_text()
                                if "yes" in talkMSG1.lower():
                                    self.SpeakText(
                                        "The mail body contains the following:")
                                    self.SpeakText(body)
                                else:
                                    self.SpeakText("You chose NO")

                            # MAIL WITH ATTACHMENT
                            elif "attachment" in content_disposition:
                                self.SpeakText(
                                    "The mail contains attachment, the contents of which will be saved in respective folders with it's name similar to that of subject of the mail")
                                filename = part.get_filename()  # download attachment
                                if filename:
                                    folder_name = self.clean(subject)
                                    if not os.path.isdir(folder_name):
                                        # make a folder for this email (named after the subject)
                                        os.mkdir(folder_name)
                                    filepath = os.path.join(folder_name, filename)
                                    # download attachment and save it
                                    open(filepath, "wb").write(
                                        part.get_payload(decode=True))

                    # NOT MULTIPART
                    else:
                        content_type = msg.get_content_type()    # extract content type of email
                        # get the email body
                        body = msg.get_payload(decode=True).decode()
                        if content_type == "text/plain":
                            self.SpeakText(
                                "Do you want to listen to the text content of the mail ? Please say YES or NO.")
                            talkMSG2 = self.speech_to_text()
                            if "yes" in talkMSG2.lower():
                                self.SpeakText("The mail body contains the following:")
                                self.SpeakText(body)
                            else:
                                self.SpeakText("You chose NO")

                    # HTML CONTENTS
                    if content_type == "text/html":
                        self.SpeakText("The mail contains an HTML part, the contents of which will be saved in respective folders with it's name similar to that of subject of the mail. You can view the html files in any browser, simply by clicking on them.")
                        # if it's HTML, create a new HTML file
                        folder_name = self.clean(subject)
                        if not os.path.isdir(folder_name):
                            # make a folder for this email (named after the subject)
                            os.mkdir(folder_name)
                        filename = "index.html"
                        filepath = os.path.join(folder_name, filename)
                        # write the file
                        open(filepath, "w").write(body)

                        # webbrowser.open(filepath)     # open in the default browser

                    self.SpeakText(f"\nEnd of message {msgCount}:")
                    msgCount += 1
                    print("="*100)
        imap.close()
        imap.logout()


    def searchMail(self):
        """
        Search mails by subject / author mail ID

        Returns:
            None: None
        """
        M = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        M.login(EMAIL_ID, PASSWORD)

        mailBoxTarget = "INBOX"
        self.SpeakText("Where do you want to search ? Say 1 for Inbox. Say 2 for Sent Mailbox. Say 3 for Drafts. Say 4 for important mails. Say 5 for Spam. Say 6 for Starred Mails. Say 7 for Bin.")
        cmb = self.speech_to_text()
        cmb = cmb.split()[-1]
        if fuzz.partial_ratio("1", cmb) >= 60 or fuzz.ratio("one", cmb) >= 60:
            mailBoxTarget = "INBOX"
            driver.get(localhost_url + "mailinInbox")
            self.SpeakText("Inbox selected.")
        elif fuzz.partial_ratio("2", cmb) >= 60 or fuzz.ratio("two", cmb) >= 60 or cmb.lower() == "tu":
            mailBoxTarget = '"[Gmail]/Sent Mail"'
            driver.get(localhost_url + "mailinSentMailbox")
            self.SpeakText("Sent Mailbox selected.")
        elif fuzz.partial_ratio("3", cmb) >= 60 or fuzz.ratio("three", cmb) >= 60:
            mailBoxTarget = '"[Gmail]/Drafts"'
            driver.get(localhost_url + "mailinDrafts")
            self.SpeakText("Drafts selected.")
        elif fuzz.partial_ratio("4", cmb) >= 60 or fuzz.ratio("four", cmb) >= 60:
            mailBoxTarget = '"[Gmail]/Important"'
            driver.get(localhost_url + "mailinImportantMails")
            self.SpeakText("Important Mails selected.")
        elif fuzz.partial_ratio("5", cmb) >= 60 or fuzz.ratio("five", cmb) >= 60:
            mailBoxTarget = '"[Gmail]/Spam"'
            driver.get(localhost_url + "mailinSpam")
            self.SpeakText("Spam selected.")
        elif fuzz.partial_ratio("6", cmb) >= 60 or fuzz.ratio("six", cmb) >= 60:
            mailBoxTarget = '"[Gmail]/Starred"'
            driver.get(localhost_url + "mailinStarredMails")
            self.SpeakText("Starred Mails selected.")
        elif fuzz.partial_ratio("7", cmb) >= 60 or fuzz.ratio("seven", cmb) >= 60:
            mailBoxTarget = '"[Gmail]/Bin"'
            driver.get(localhost_url + "mailinBin")
            self.SpeakText("Bin selected.")
        else:
            driver.get(localhost_url + "mailinInbox")
            self.SpeakText("Wrong choice. Hence, default option Inbox wil be selected.")
        
        
        M.select(mailBoxTarget)

        self.SpeakText("Say 1 to search mails from a specific sender. Say 2 to search mail with respect to the subject of the mail.")
        mailSearchChoice = self.speech_to_text()
        mailSearchChoice = mailSearchChoice.split()[-1]
        if mailSearchChoice == "1" or mailSearchChoice.lower() == "one":
            self.SpeakText("Please mention the sender email ID you want to search.")
            searchSub = self.speech_to_text()
            searchSub = searchSub + '@gmail.com'
            #searchSub = searchSub.replace("at the rate", "@")
            searchSub = searchSub.replace(" ", "")
            status, messages = M.search(None, f'FROM "{searchSub}"')
        elif mailSearchChoice == "2" or mailSearchChoice.lower() == "two" or mailSearchChoice.lower() == "tu":
            self.SpeakText("Please mention the subject of the mail you want to search.")
            searchSub = self.speech_to_text()
            status, messages = M.search(None, f'SUBJECT "{searchSub}"')
        else:
            self.SpeakText(
                "Wrong choice. Performing default operation. Please mention the subject of the mail you want to search.")
            searchSub = self.speech_to_text()
            status, messages = M.search(None, f'SUBJECT "{searchSub}"')

        if str(messages[0]) == "b''":
            self.SpeakText(f"Mail not found in {mailBoxTarget}.")
            return None

        msgCount = 1
        for i in messages:
            self.SpeakText(f"Message {msgCount}:")
            res, msg = M.fetch(i, "(RFC822)")   # fetch the email message by ID
            for response in msg:
                if isinstance(response, tuple):
                    # parse a bytes email into a message object
                    msg = email.message_from_bytes(response[1])

                    subject, encoding = decode_header(msg["Subject"])[
                        0]    # decode the email subject
                    if isinstance(subject, bytes):
                        # if it's a bytes, decode to str
                        subject = subject.decode(encoding)

                    From, encoding = decode_header(msg.get("From"))[
                        0]      # decode email sender
                    if isinstance(From, bytes):
                        From = From.decode(encoding)
                    self.SpeakText("Subject: " + subject)
                    FromArr = From.split()
                    FromName = " ".join(namechar for namechar in FromArr[0:-1])
                    self.SpeakText("From: " + FromName)
                    self.SpeakText("Sender mail: " + FromArr[-1])

                    # MULTIPART
                    if msg.is_multipart():
                        for part in msg.walk():  # iterate over email parts
                            content_type = part.get_content_type()      # extract content type of email
                            content_disposition = str(
                                part.get("Content-Disposition"))
                            try:
                                body = part.get_payload(
                                    decode=True).decode()   # get the email body
                            except:
                                pass

                            # PLAIN TEXT MAIL
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                self.SpeakText(
                                    "Do you want to listen to the text content of the mail ? Please say YES or NO.")
                                talkMSG1 = self.speech_to_text()
                                if "yes" in talkMSG1.lower():
                                    self.SpeakText(
                                        "The mail body contains the following:")
                                    self.SpeakText(body)
                                else:
                                    self.SpeakText("You chose NO")

                            # MAIL WITH ATTACHMENT
                            elif "attachment" in content_disposition:
                                self.SpeakText(
                                    "The mail contains attachment, the contents of which will be saved in respective folders with it's name similar to that of subject of the mail")
                                filename = part.get_filename()  # download attachment
                                if filename:
                                    folder_name = self.clean(subject)
                                    if not os.path.isdir(folder_name):
                                        # make a folder for this email (named after the subject)
                                        os.mkdir(folder_name)
                                    filepath = os.path.join(folder_name, filename)
                                    # download attachment and save it
                                    open(filepath, "wb").write(
                                        part.get_payload(decode=True))

                    # NOT MULTIPART
                    else:
                        content_type = msg.get_content_type()    # extract content type of email
                        # get the email body
                        body = msg.get_payload(decode=True).decode()
                        if content_type == "text/plain":
                            self.SpeakText(
                                "Do you want to listen to the text content of the mail ? Please say YES or NO.")
                            talkMSG2 = self.speech_to_text()
                            if "yes" in talkMSG2.lower():
                                self.SpeakText("The mail body contains the following:")
                                self.SpeakText(body)
                            else:
                                self.SpeakText("You chose NO")

                    # HTML CONTENTS
                    if content_type == "text/html":
                        self.SpeakText("The mail contains an HTML part, the contents of which will be saved in respective folders with it's name similar to that of subject of the mail. You can view the html files in any browser, simply by clicking on them.")
                        # if it's HTML, create a new HTML file
                        folder_name = self.clean(subject)
                        if not os.path.isdir(folder_name):
                            # make a folder for this email (named after the subject)
                            os.mkdir(folder_name)
                        filename = "index.html"
                        filepath = os.path.join(folder_name, filename)
                        # write the file
                        open(filepath, "w").write(body)

                        # webbrowser.open(filepath)     # open in the default browser

                    self.SpeakText(f"\nEnd of message {msgCount}:")
                    msgCount += 1
                    print("="*100)

        M.close()
        M.logout()

# Creating Instances
if __name__ == "__main__":
    speech = SpeechMail()
    driver.get("http://127.0.0.1:5000/")
    if EMAIL_ID != "" and PASSWORD != "":
        speech.SpeakText("Say 1 to send a mail. Say 2 to get your mailbox status. Say 3 to search a mail. Say 4 to get the last 3 mails.")
        choice = speech.speech_to_text()
        # choice = choice.split()[-1]
        if fuzz.partial_ratio("1", choice) >= 60 or fuzz.ratio("one", choice) >= 60:
        # if choice == '1' or choice == 'one':
            driver.get("http://127.0.0.1:5000/composeMail")
            speech.composeMail()
        elif fuzz.partial_ratio("2", choice) >= 60 or fuzz.ratio("two", choice) >= 60:
        # elif choice == '2' or choice == 'too' or choice == 'two' or choice == 'to' or choice == 'tu' or choice.lower() == 'TW'.lower():
            driver.get("http://127.0.0.1:5000/MailBoxStatus")
            speech.getMailBoxStatus()
        elif fuzz.partial_ratio("3", choice) >= 60 or fuzz.ratio("three", choice) >= 60:
        # elif choice == '3' or choice == 'tree' or choice == 'three':
            driver.get("http://127.0.0.1:5000/Mail")
            speech.searchMail()
        elif fuzz.partial_ratio("4", choice) >= 60 or fuzz.ratio("four", choice) >= 60:
        # elif choice == '4' or choice == 'four' or choice == 'for':
            driver.get("http://127.0.0.1:5000/Mail")
            speech.getLatestMails()
        else:
            speech.SpeakText("Wrong choice. Please say only the number")
    else:
        speech.SpeakText("Both Email ID and Password should be present")
        