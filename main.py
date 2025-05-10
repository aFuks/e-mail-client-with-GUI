import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import imaplib
import email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from PIL import Image, ImageTk
from io import BytesIO

IMAP_SERVER = 'imap.gmail.com'
IMAP_PORT = 000
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 000
EMAIL_ADDRESS = '***'
PASSWORD = '***'

class EmailViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Skrzynka pocztowa")

        self.create_search_bar() 

        self.table_frame = tk.Frame(root)
        self.table_frame.pack(padx=10, pady=10)

        self.create_table()

        self.button_frame = tk.Frame(root)
        self.button_frame.pack(pady=5)

        self.open_button = tk.Button(self.button_frame, text="Otwórz", command=self.open_selected_email)
        self.open_button.pack(side="left", padx=5)

        self.refresh_button = tk.Button(self.button_frame, text="Odśwież", command=self.fetch_last_10_emails)
        self.refresh_button.pack(side="left", padx=5)

        self.compose_button = tk.Button(self.button_frame, text="Napisz", command=self.compose_email)
        self.compose_button.pack(side="left", padx=5)

        self.emails = []
        self.fetch_last_10_emails()
        
        self.monitor_incoming_emails()

    def create_search_bar(self):
        self.search_frame = tk.Frame(self.root)
        self.search_frame.pack(pady=5)

        self.search_label = tk.Label(self.search_frame, text="Wyszukaj:")
        self.search_label.pack(side="left", padx=5)

        self.search_entry = tk.Entry(self.search_frame, width=30)
        self.search_entry.pack(side="left", padx=5)

        self.search_button = tk.Button(self.search_frame, text="Szukaj", command=self.search_emails)
        self.search_button.pack(side="left", padx=5)

    def fetch_last_10_emails(self):
        self.table.delete(*self.table.get_children())
        try:
            mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
            mail.login(EMAIL_ADDRESS, PASSWORD)
            mail.select('inbox')
            result, data = mail.search(None, 'ALL')
            email_ids = data[0].split()
            email_ids = email_ids[-10:]  # only last ten emails are downloaded

            self.emails = []
            self.email_index_offset = int(email_ids[-1]) - 9  

            keyword = self.search_entry.get().lower()  # search by keyword

            for email_id in email_ids:
                result, email_data = mail.fetch(email_id, '(RFC822)')
                raw_email = email_data[0][1]
                msg = email.message_from_bytes(raw_email)
                sender = msg['From']
                subject = msg['Subject']
                body = ""
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8")
                        break
                
                if not keyword or keyword in subject.lower() or keyword in sender.lower() or keyword in body.lower():
                    self.emails.append((sender, subject))
                    self.table.insert("", "end", text=email_id, values=(sender, subject))

            mail.close()
            mail.logout()

        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił błąd: {str(e)}")

    def display_emails(self, emails):
        self.emails = emails
        self.email_index_offset = 0

        for idx, email_info in enumerate(emails):
            sender, subject = email_info
            self.table.insert("", "end", text=idx, values=(sender, subject))

    def search_emails(self):
        keyword = self.search_entry.get().lower()
        self.table.delete(*self.table.get_children())
        try:
            mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
            mail.login(EMAIL_ADDRESS, PASSWORD)
            mail.select('inbox')
            result, data = mail.search(None, 'ALL')
            email_ids = data[0].split()
            email_ids = email_ids[-10:]  

            self.emails = []
            self.email_index_offset = int(email_ids[-1]) - 9 

            for email_id in email_ids:
                result, email_data = mail.fetch(email_id, '(RFC822)')
                raw_email = email_data[0][1]
                msg = email.message_from_bytes(raw_email)
                sender = msg['From']
                subject = msg['Subject']
                body = ""
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8")
                        break
                
                if keyword in subject.lower() or keyword in sender.lower() or keyword in body.lower():
                    self.emails.append((sender, subject))
                    self.table.insert("", "end", text=email_id, values=(sender, subject))

            mail.close()
            mail.logout()

        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił błąd: {str(e)}")

    def create_table(self):
        self.table = ttk.Treeview(self.table_frame, columns=("Sender", "Subject"))
        self.table.heading("#0", text="ID")
        self.table.heading("Sender", text="Nadawca")
        self.table.heading("Subject", text="Temat")
        self.table.pack(expand=True, fill="both")

    def open_selected_email(self):
        selected_item = self.table.selection()
        if selected_item:
            email_id = self.table.item(selected_item, "text")
            idx = int(email_id) - self.email_index_offset
            if 0 <= idx < len(self.emails):
                self.show_email_content(idx)
            else:
                messagebox.showwarning("Błąd", "Wybrany e-mail nie istnieje.")
        else:
            messagebox.showwarning("Błąd", "Nie wybrano żadnego e-maila.")

    def show_email_content(self, idx):
        selected_item = self.table.selection()
        if selected_item:
            email_id = self.table.item(selected_item, "text")
            if email_id:
                email_id = int(email_id)
                try:
                    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
                    mail.login(EMAIL_ADDRESS, PASSWORD)
                    mail.select('inbox')
                    result, data = mail.search(None, 'ALL')
                    email_ids = data[0].split()
                    email_id = email_ids[-10:][email_id - self.email_index_offset] 

                    result, email_data = mail.fetch(email_id, '(RFC822)')
                    raw_email = email_data[0][1]
                    msg = email.message_from_bytes(raw_email)

                    sender = msg['From']
                    subject = msg['Subject']
                    body = ""
                    images = []

                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True).decode("utf-8")
                        elif part.get_content_type().startswith("image"):
                            image_data = part.get_payload(decode=True)
                            image = Image.open(BytesIO(image_data))
                            images.append(image)

                    self.display_email_content(sender, subject, body, images)

                    mail.close()
                    mail.logout()
                except Exception as e:
                    messagebox.showerror("Błąd", f"Wystąpił błąd: {str(e)}")
            else:
                messagebox.showwarning("Błąd", "Wybrany e-mail nie istnieje.")
        else:
            messagebox.showwarning("Błąd", "Nie wybrano żadnego e-maila.")

    def display_email_content(self, sender, subject, body, images):
        self.email_content_window = tk.Toplevel(self.root)
        self.email_content_window.title("Treść maila")

        sender_label = tk.Label(self.email_content_window, text="From: " + sender)
        subject_label = tk.Label(self.email_content_window, text="Subject: " + subject)
        body_text = scrolledtext.ScrolledText(self.email_content_window, wrap=tk.WORD)
        body_text.insert(tk.END, "Body: \n" + body)
        body_text.config(state=tk.DISABLED)

        sender_label.pack()
        subject_label.pack()
        body_text.pack(expand=True, fill="both")

        for image in images:
            tk_image = ImageTk.PhotoImage(image)
            image_label = tk.Label(self.email_content_window, image=tk_image)
            image_label.image = tk_image
            image_label.pack()

    def compose_email(self):
        self.compose_window = tk.Toplevel(self.root)
        self.compose_window.title("Napisz e-mail")

        self.to_label = tk.Label(self.compose_window, text="To:")
        self.to_entry = tk.Entry(self.compose_window)
        self.to_label.grid(row=0, column=0, sticky="e")
        self.to_entry.grid(row=0, column=1, padx=5, pady=5)

        self.subject_label = tk.Label(self.compose_window, text="Subject:")
        self.subject_entry = tk.Entry(self.compose_window)
        self.subject_label.grid(row=1, column=0, sticky="e")
        self.subject_entry.grid(row=1, column=1, padx=5, pady=5)

        self.body_label = tk.Label(self.compose_window, text="Body:")
        self.body_text = scrolledtext.ScrolledText(self.compose_window, wrap=tk.WORD)
        self.body_label.grid(row=2, column=0, sticky="ne")
        self.body_text.grid(row=2, column=1, padx=5, pady=5)

        self.send_button = tk.Button(self.compose_window, text="Wyślij", command=self.send_email)
        self.send_button.grid(row=3, columnspan=2, pady=5)

    def send_email(self):
        to_address = self.to_entry.get()
        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", tk.END)

        try:
            msg = MIMEMultipart()
            msg['From'] = EMAIL_ADDRESS
            msg['To'] = to_address
            msg['Subject'] = subject

            msg.attach(MIMEText(body, 'plain'))

            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(EMAIL_ADDRESS, PASSWORD)
            text = msg.as_string()
            server.sendmail(EMAIL_ADDRESS, to_address, text)
            server.quit()

            messagebox.showinfo("Sukces", "E-mail został wysłany pomyślnie.")

            self.compose_window.destroy()

        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił błąd podczas wysyłania e-maila: {str(e)}")

    def monitor_incoming_emails(self):
        try:
            mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
            mail.login(EMAIL_ADDRESS, PASSWORD)
            mail.select('inbox')

            mail.search(None, 'Unseen')
            new_emails = mail.search(None, 'Unseen')[1][0].split()

            for email_id in new_emails:
                result, email_data = mail.fetch(email_id, '(RFC822)')
                raw_email = email_data[0][1]
                msg = email.message_from_bytes(raw_email)

                # send auto respose
                self.send_automatic_response(msg)

            mail.close()
            mail.logout()

        except Exception as e:
            print(f"Wystąpił błąd podczas monitorowania nowych e-maili: {str(e)}")

        # refresh after 10sec
        self.root.after(10000, self.monitor_incoming_emails)

    def send_automatic_response(self, received_msg):
        sender = received_msg['From']

        response_subject = "Automatyczna odpowiedź: Dziękuję za kontakt"
        response_body = "Dziękuję za kontakt, odezwę się jak najszybciej."

        try:
            msg = MIMEMultipart()
            msg['From'] = EMAIL_ADDRESS
            msg['To'] = sender
            msg['Subject'] = response_subject

            msg.attach(MIMEText(response_body, 'plain'))

            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(EMAIL_ADDRESS, PASSWORD)
            text = msg.as_string()
            server.sendmail(EMAIL_ADDRESS, sender, text)
            server.quit()

        except Exception as e:
            print(f"Wystąpił błąd podczas wysyłania automatycznej odpowiedzi: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailViewer(root)
    root.mainloop()
