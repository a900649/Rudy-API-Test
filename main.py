from waitress import serve
from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
import sql_function
import my_function
from azure.storage.blob import BlobServiceClient
import pandas as pd
from datetime import datetime,timezone,timedelta
import variables as v
import json
from io import BytesIO
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib, ssl

app=Flask(__name__)
app.config['JSON_AS_ASCII'] = False
CORS(app)

@app.route("/",methods = ["GET"])
def home():

    return {"Status":"OK"}

@app.route("/authenticate",methods = ['POST'])
def authenticate():

    parameter_dict = request.get_json()

    account = parameter_dict["Account"]
    password = parameter_dict["Password"]
    token = parameter_dict["Token"]

    status,data,message = sql_function.authenticate_user(account,password,token)

    return_dict = {"Status":status,"Data":data,"Message":message}

    return return_dict

@app.route("/edit_data",methods = ['POST'])
def edit_data():

    parameter_dict = request.get_json()

    row_id = parameter_dict["RowID"]
    action = parameter_dict["Action"]
    user_id = parameter_dict["User ID"]
    email = parameter_dict["Email"]
    title = parameter_dict["Title"]
    application_date_str = parameter_dict["Application Date"]
    occurrence_date_str = parameter_dict["Occurrence Date"]
    total_amount = parameter_dict["Total Amount"]
    description = parameter_dict["Description"]
    toll_amount = parameter_dict["Toll Amount"]
    parking_amount = parameter_dict["Parking Amount"]
    petrol_amount = parameter_dict["Petrol Amount"]
    car_allowance_amount = parameter_dict["Car Allowance Amount"]
    printing_stationery_amount = parameter_dict["Printing Stationery Amount"]
    accomodation_amount = parameter_dict["Accomodation Amount"]
    meal_allowance_amount = parameter_dict["Meal Allowance Amount"]
    shift_allowance_amount = parameter_dict["Shift Allowance Amount"]
    travelling_fee_amount = parameter_dict["Travelling Fee Amount"]
    medical_fee_amount = parameter_dict["Medical Fee Amount"]
    gift_promotion_amount = parameter_dict["Gift Promotion Amount"]
    courier_charges_amount = parameter_dict["Courier Charges Amount"]
    research_development_amount = parameter_dict["Research Development Amount"]
    mv_maintenance_amount = parameter_dict["MV Maintenance Amount"]
    unloading_fee_sugar_amount = parameter_dict["Unloading Fee-Sugar Amount"]
    subscriotion_fee_amount = parameter_dict["Subscriotion Fee Amount"]

    approve_date_str = parameter_dict["Approve Date"]
    account_date_str = parameter_dict["Account Date"]

    status, row_id = sql_function.edit_data(row_id, action, application_date_str, occurrence_date_str, user_id, email, title, total_amount, description,
                                            toll_amount, parking_amount, petrol_amount,car_allowance_amount ,
                                            printing_stationery_amount, accomodation_amount, meal_allowance_amount, shift_allowance_amount, travelling_fee_amount,
                                            medical_fee_amount, gift_promotion_amount, courier_charges_amount, research_development_amount, mv_maintenance_amount,
                                            unloading_fee_sugar_amount, subscriotion_fee_amount, approve_date_str, account_date_str)


    return_dict = {"Status":status,"Row ID":row_id}

    return return_dict

@app.route("/upload_file/<row_id>",methods = ['POST'])
def upload_file(row_id):
    print(row_id)
    blob_service_client = BlobServiceClient.from_connection_string(v.blob_connection_string)
    container_client = blob_service_client.get_container_client(v.blob_container)
    url_list = []
    for i in range(1,49):
        file = request.files.get('image{}'.format(i))
        if file != None:
            filename = row_id + "_" + str(i) + "." + file.filename.split(".")[-1]
            blob_client = container_client.get_blob_client(filename)
            blob_client.upload_blob(file, overwrite=True)
            url_list.append(v.blob_url + filename)
        else:
            url_list.append(None)

    status = sql_function.edit_attachment_str(row_id, url_list)


    return {"Status":status}


@app.route("/get_data_by_userid",methods = ['POST'])
def get_data_by_userid():

    parameter_dict = request.get_json()
    start_date_str = parameter_dict["Start Date"]
    end_date_str = parameter_dict["End Date"]
    user_id = parameter_dict["User ID"]
    schedule_type_str = parameter_dict["Schedule Type"]

    schedule_type_str_list = schedule_type_str.split("@")
    schedule_type_list = []
    if schedule_type_str_list[0].lower() == "true":
        schedule_type_list.append("Waiting")
    if schedule_type_str_list[1].lower() == "true":
        schedule_type_list.append("Approved")
        schedule_type_list.append("Rejected")

    status, data_list, col_name_list = sql_function.get_data_by_userid(start_date_str, end_date_str, [user_id], schedule_type_list)

    return_dict = {}
    if status == True:
        for i in range(0,len(col_name_list)):
            col_data = []
            col_name = col_name_list[i][0]
            for r in range(0,len(data_list)):
                data = data_list[r][i]
                # 調整型態
                if col_name in ["Application Date",'Occurrence Date']:
                    data = datetime.strftime(data, '%d/%m/%Y')
                if col_name == "Total Amount":
                    data = float(data)
                if col_name == "Amount Detail":
                    data = json.loads(data)
                if col_name == "URL Detail":
                    data = json.loads(data)

                col_data.append(data)

            return_dict[col_name] = col_data

    return {"Status":status,"Data":return_dict}


@app.route("/get_approval_data",methods = ['POST'])
def get_approval_data():

    parameter_dict = request.get_json()
    start_date_str = parameter_dict["Start Date"]
    end_date_str = parameter_dict["End Date"]
    email = parameter_dict["Email"]
    schedule_type_str = parameter_dict["Schedule Type"]

    schedule_type_str_list = schedule_type_str.split("@")
    schedule_type_list = []
    if schedule_type_str_list[0].lower() == "true":
        schedule_type_list.append("Waiting")
    if schedule_type_str_list[1].lower() == "true":
        schedule_type_list.append("Approved")
        schedule_type_list.append("Rejected")

    status, data_list, col_name_list = sql_function.get_my_approval_data(start_date_str, end_date_str, email, schedule_type_list)

    return_dict = {}
    if status == True:
        for i in range(0, len(col_name_list)):
            col_data = []
            col_name = col_name_list[i][0]
            for r in range(0, len(data_list)):
                data = data_list[r][i]
                # 調整型態
                if col_name in ["Application Date", 'Occurrence Date']:
                    data = datetime.strftime(data, '%d/%m/%Y')
                if col_name == "Total Amount":
                    data = float(data)
                if col_name == "Amount Detail":
                    data = json.loads(data)
                if col_name == "URL Detail":
                    data = json.loads(data)

                col_data.append(data)

            return_dict[col_name] = col_data

    return {"Status": status, "Data": return_dict}

@app.route("/alter_schedule",methods = ['POST'])
def alter_schedule():

    parameter_dict = request.get_json()
    data_type = parameter_dict["Data Type"]
    row_id = parameter_dict["Row ID"]
    schedule_type = parameter_dict["Schedule Type"]

    status = sql_function.alter_schedule(data_type, row_id, schedule_type)

    return {"Status": status}

@app.route("/upload_excel",methods = ['POST'])
def upload_excel():

    print("enter upload_excel")
    file = request.files.get('file_1')
    a = pd.read_excel(file)
    print(a)
    return {"Status":"OK"}


@app.route("/send_mail_post",methods = ["POST"])
def send_mail_post():

    smtp_server = "smtp.office365.com"
    port = 587
    sender = "paultest20240617@outlook.com"
    mail_user = "paultest20240617@outlook.com"
    mail_password = "FlaskAPITest@"

    convert_string = "@1@2@3@$$#"

    def byte_string_convert(byte_string):
        return_string = byte_string.decode("utf-8").replace("\n", convert_string).replace(","+convert_string,",").replace("{"+convert_string,"{").replace(convert_string+"}","}")
        return return_string

    parameter_dict = json.loads(byte_string_convert(request.get_data()))

    subject = parameter_dict["Subject"].replace(convert_string,"\n")
    content = parameter_dict["Content"].replace(convert_string,"\n")
    receiver = parameter_dict["Receiver"].replace(convert_string,"\n")

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

if __name__ == '__main__':
    # app.run(host="127.0.0.1", port=8080, debug=True)
    serve(app, host="0.0.0.0", port=8080, threads=100,backlog=9216,connection_limit=10000)