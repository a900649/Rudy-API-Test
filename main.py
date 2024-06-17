from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib, ssl
import pandas as pd
from io import BytesIO
from flask import Flask, send_from_directory, request, make_response, send_file
import variable as v

app = Flask(__name__)

@app.route("/",methods = ["GET"])
def home():
    return {"Status": "OK"}

@app.route("/send_mail_get/<subject>/<content>/<receiver>",methods = ["GET"])
def send_mail_get(subject,content,receiver):

    smtp_server = v.smtp_server
    port = v.port
    sender = v.sender
    mail_user = v.mail_user
    mail_password = v.mail_password

    server = smtplib.SMTP(smtp_server, port)

    # 建立連線
    context = ssl.create_default_context()
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['FROM'] = sender
    msg['To'] = receiver

    # 建立內文文字
    att2 = MIMEText(content)
    msg.attach(att2)

    server.starttls(context=context)
    server.login(mail_user, mail_password)
    server.sendmail(sender, receiver, msg.as_string())

    return {"Status": "OK"}

@app.route("/send_mail_post",methods = ["POST"])
def send_mail_post(subject,content,receiver):

    smtp_server = v.smtp_server
    port = v.port
    sender = v.sender
    mail_user = v.mail_user
    mail_password = v.mail_password

    parameter_dict = request.get_json()
    subject = parameter_dict["Subject"]
    content = parameter_dict["Content"]
    receiver = parameter_dict["Receiver"]

    server = smtplib.SMTP(smtp_server, port)

    # 建立連線
    context = ssl.create_default_context()
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['FROM'] = sender
    msg['To'] = receiver

    # 建立內文文字
    att2 = MIMEText(content)
    msg.attach(att2)

    server.starttls(context=context)
    server.login(mail_user, mail_password)
    server.sendmail(sender, receiver, msg.as_string())

    return {"Status": "OK"}


@app.route("/download_type_1",methods = ["GET","POST"])
def download_type_1():
    df = pd.DataFrame([[1, 2], [3, 4]], columns=["A", "B"])

    out = BytesIO()
    writer = pd.ExcelWriter(out, engine='xlsxwriter')
    df.to_excel(excel_writer=writer, sheet_name='Sheet1',)
    writer.close()

    out.seek(0)

    rv = make_response(out.getvalue())
    out.close()
    rv.headers['Content-Type'] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    rv.headers["Cache-Control"] = "no-cache"
    rv.headers['Content-Disposition'] = 'attachment; filename={}.xlsx'.format('Output')

    return rv


@app.route("/download_type_2",methods = ["GET","POST"])
def download_type_2():
    df = pd.DataFrame([[1, 2], [3, 4]], columns=["A", "B"])

    out = BytesIO()
    writer = pd.ExcelWriter(out, engine='xlsxwriter')
    df.to_excel(excel_writer=writer, sheet_name='Sheet1',)
    writer.close()

    out.seek(0)


    return send_file(out, download_name="output.xlsx",mimetype="application/vnd.ms-excel", as_attachment=True)

if __name__ == '__main__':
    # app.run(host="0.0.0.0", port=8000)

    app.run(host="127.0.0.1", port=8060)
