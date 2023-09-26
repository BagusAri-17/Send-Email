import os
import smtplib, ssl
import time
import pandas as pd
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from config import username_email, password_email

receivers = pd.read_excel("receivers.xlsx", usecols="A, B, C")

image_folder = "./tiket-g-kekirim"

# send email
context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(username_email, password_email)
    for i in receivers.index:
        msg = MIMEMultipart()
        msg["subject"] = "D-DAY WEBINAR NASIONAL INFORMATIKA 2023"
        msg["From"] = "WEBINAR NASIONAL INFORMATIKA 2023"
        body = (
            """
        <!DOCTYPE html>
        <html lang="en">

        <head>
            <meta charset="UTF-8">
            <meta http-equiv="X-UA-Compatible" content="IE=edge">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Automata Tiket Webinar</title>

            <style>
                @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@100;300;400;500;700;900&display=swap');
                @media (prefers-color-scheme: dark){
                    .section p , .section small{
                        color: black !important;
                    }
                }

                * {
                    font-family: 'Roboto', sans-serif;
                    box-sizing: border-box;
                    padding: 0;
                    margin: 0;
                }

                .container {
                    height: 100vh;
                }

                .header .text {
                    color: white;
                    padding-left: 1.55rem;
                }

                .header,
                .footer {
                    background: #dcdce4;
                    background: linear-gradient(90deg, #282624, #8a761a 100%, #00d4ff 0);
                    color: white;
                }

                .section {
                    /*background: url("https://d33wubrfki0l68.cloudfront.net/static/media/fb1f349208f1d6f59c9a196fdb5dc23cabe80b4e/bg-web.4e83223833905a64c54b.jpg");*/
                    text-align: justify;
                    line-height: 1.5rem;
                }

                .btn-a {
                    padding: 0.8rem 5rem;
                    background-color: #6f5f14;
                    color: white !important;
                    border: none;
                    margin: auto !important;
                    text-decoration: none;
                    border-radius: 30px;
                    filter: drop-shadow(4px 4px 4px rgba(0, 0, 0, 0.25));
                    font-weight: 700;
                }

                .btn-a:hover{
                    opacity: 90%;
                }

                .section-header {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 1rem;
                }

                .section-header small {
                    color: gray;
                }

                .footer {
                    text-align: center;
                    color: white;
                    padding-top: 0.5rem;
                    padding-bottom: 0.5rem;
                }
                p{
                    color: black !important;
                }
            </style>
        </head>

        <body>
            <div class="container">
                <div class="child">
                    <table cellspading='0' cellpading='0' border='0' class='header' width='400px'
                        style='padding: 1rem 3rem; border-left: 2px solid #090979; border-right: 2px solid #090979; border-top-left-radius: 10px !important; border-top-right-radius: 10px !important;'>
                        <tr style="align-items: center;">
                            <td width='90px'>
                                <img src="https://raw.githubusercontent.com/BagusAri-17/Bot-WA/main/logowebnas.png"
                                    width="100%" class='ps-4 m-0 p-0 my-auto'>
                            </td>
                            <td width='210px' style='padding-left: 1.5rem;'>
                                <h2>WEBINAR NASIONAL INFORMATIKA</h2>
                                <small>Webinar Nasional</small>
                            </td>
                        </tr>
                    </table>
                    <table class='section' cellspading='0' cellpading='0' border='0' style='padding: 2rem 1.5rem; border-left: 2px solid #090979; border-right: 2px solid #090979;'
                        width='400px'>
                        <tr>
                            <td colspan="2" class="bodyy">
                                <p>
                                    Halo Seluruh Civitas Indonesia!🙌
                                </p>
                                <p style='padding-top: 0.5rem;'>
                                    Tidak terasa nih Webinar Nasional Informatika 2023 akan segera hadir. Persiapkan dirimu yaa pada:
                                </p>
                                <p style='padding-top: 0.5rem;'>‼SAVE THE DATE‼</p>
                                <p style='padding-top: 0.5rem; text-align: left;'>
                                    📆: <b>Minggu, 24 September 2023</b><br>
                                    🕓: <b>09.00 WITA - selesai</b><br>
                                    📍: <b>Zoom Meeting (link akan dikirimkan melalui grup peserta)</b><br>
                                </p>
                                <p style="padding-top: 0.5rem;">
                                    🚀 Jangan sampai ketinggalan berita terbaru seputar Webinar Nasional Informatika 2023! Bergabunglah dengan kami di Grup Telegram resmi kami.
                                </p>
                                <p style="padding-top: 0.5rem; text-align: left;">
                                    📣 <b>Grup Telegram Webinar Nasional Informatika 2023:</b><br>
                                    🔗: <a href="https://bit.ly/GrupTelegramWebnasInformatika2023">https://bit.ly/GrupTelegramWebnasInformatika2023</a><br>
                                </p>
                                <p style="padding-top: 0.5rem; text-align: left;">
                                    Untuk Peserta yang akan join ke grup telegram, diharapkan untuk mengubah nama akun menggunakan format sebagai berikut:<br>
                                    <b>Nomor Tiket_Nama Pendaftar</b> (Contoh: <b>1_Ardiska</b>)<br>
                                </p>
                                <p style="padding-top: 0.5rem; text-align: left;">
                                    Nomor Tiket Kamu: <b>"""
            + str(receivers["nomor"][i])
            + """</b><br>
                                    Tata Tertib Peserta: <a href="https://bit.ly/TataTertibWebnasInformatika2023">https://bit.ly/TataTertibWebnasInformatika2023</a><br>
                                    Susunan Acara: <a href="https://bit.ly/SusunanAcaraPesertaWebinarNasionalInformatika2023">https://bit.ly/SusunanAcaraPesertaWebinarNasionalInformatika2023</a><br>
                                </p>
                                <p style="padding-top: 0.5rem; text-align: left;">
                                    Dimohon kepada seluruh peserta membaca dan menaati tata tertib tersebut yaa. Terima kasih✨<br>
                                </p>
                                <p style="padding-top: 0.5rem; text-align: left;">
                                    Kami tunggu kehadiranmu, see you!😉🤩<br>
                                    ======================<br>
                                    Webinar Nasional Informatika 2023!<br>
                                </p>
                                <div style="padding-top: 2rem; text-align: center; width: 100%">
                                    <small style = 'color: black !important;'>
                                        <b>
                                            Terdapat kendala? Silakan hubungi contact person berikut.
                                        </b>
                                    </small>
                                </div>
                                <div style='text-align: center; margin-top: 1rem;'>
                                    <a href='https://ig.me/m/webnas.informatika' class='btn-a'>
                                        Contact Person
                                    </a>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table cellspading='0' class = 'footer' cellpading='0' border='0' width='400px' style='border-left: 2px solid #090979; border-right: 2px solid #090979; border-bottom-left-radius: 10px; border-bottom-right-radius: 10px;'>
                        <tr>
                            <td align="center" width = '400px'>
                                <small>&copy; Webinar Nasional informatika 2023</small>
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
        </body>
        </html>
        """
        )

        msg.attach(MIMEText(body, "html"))

        img_filename = f"{receivers['nomor'][i]}.jpg"
        img_path = os.path.join(image_folder, img_filename)

        with open(img_path, "rb") as img_file:
            img_data = img_file.read()
            img_mime = MIMEImage(img_data, "jpg")
            img_mime.add_header("Content-ID", f"<{img_filename}>")
            img_mime.add_header(
                "Content-Disposition", f"inline; filename={img_filename}"
            )
            msg.attach(img_mime)

        server.sendmail(
            from_addr=username_email,
            to_addrs=receivers["email"][i],
            msg=msg.as_string(),
        )
        print(f"{receivers['nomor'][i]} - {receivers['email'][i]} send successfully")
        time.sleep(1)
