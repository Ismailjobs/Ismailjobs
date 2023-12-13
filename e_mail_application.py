import tkinter as tk
from tkinter import messagebox
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class EmailClientOutlook:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook E-posta İstemcisi")

        self.giris_label = tk.Label(root, text="E-posta Girişi")
        self.giris_label.pack()

        self.kullanici_label = tk.Label(root, text="Kullanıcı Adı:")
        self.kullanici_label.pack()
        self.kullanici_entry = tk.Entry(root)
        self.kullanici_entry.pack()

        self.sifre_label = tk.Label(root, text="Şifre:")
        self.sifre_label.pack()
        self.sifre_entry = tk.Entry(root, show="*")
        self.sifre_entry.pack()

        self.alan_label = tk.Label(root, text="Alıcı E-posta Adresi:")
        self.alan_label.pack()
        self.alan_entry = tk.Entry(root)
        self.alan_entry.pack()

        self.konu_label = tk.Label(root, text="E-posta Konusu:")
        self.konu_label.pack()
        self.konu_entry = tk.Entry(root)
        self.konu_entry.pack()

        self.icerik_label = tk.Label(root, text="E-posta İçeriği:")
        self.icerik_label.pack()
        self.icerik_text = tk.Text(root, height=10)
        self.icerik_text.pack()

        self.gonder_button = tk.Button(root, text="E-posta Gönder", command=self.eposta_gonder)
        self.gonder_button.pack()

    def eposta_gonder(self):
        kullanici = self.kullanici_entry.get()
        sifre = self.sifre_entry.get()
        alici = self.alan_entry.get()
        konu = self.konu_entry.get()
        icerik = self.icerik_text.get("1.0", tk.END)

        try:
            # E-posta gönderme işlemi
            msg = MIMEMultipart()
            msg['From'] = kullanici
            msg['To'] = alici
            msg['Subject'] = konu

            body = icerik
            msg.attach(MIMEText(body, 'plain'))

            # Outlook SMTP sunucusuna bağlanma ve gönderme
            server = smtplib.SMTP('smtp.office365.com', 587)
            server.starttls()
            server.login(kullanici, sifre)
            server.sendmail(kullanici, alici, msg.as_string())
            server.quit()

            messagebox.showinfo("Bilgi", "E-posta başarıyla gönderildi!")
        except Exception as e:
            messagebox.showerror("Hata", "E-posta gönderme hatası: " + str(e))

if __name__ == "__main__":
    root = tk.Tk()
    uygulama = EmailClientOutlook(root)
    root.mainloop()
