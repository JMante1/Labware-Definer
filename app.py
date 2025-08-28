from flask import (
    Flask,
    render_template,
    request,
    abort,
    redirect,
    make_response,
    send_from_directory,
)
import sys
import traceback
import pandas as pd
import src.mantis.mantis_labware as ml
import src.integra.integra_labware as il
import src.tecan.tecan_labware as tl
import tempfile
from datetime import datetime
import os
import sys
import zipfile


app = Flask(__name__)

@app.route("/status")
def status():
    return "The Server is running"


@app.route("/public/<file_name>")
def success(file_name):
    top_path = os.path.split(os.path.abspath(__file__))[0]
    path = os.path.join(top_path, "public")
    print(path, file_name)
    try:
        return send_from_directory(path, file_name)
    except:
        with open(
            os.path.join(top_path, "templates", "Static_File_Not_Found.html")
        ) as file:
            error_message = file.read()

        error_message = error_message.replace("REPLACE_FILENAME", file_name)
        return error_message, 404


@app.route("/home")
def home():
    return render_template("home.html")


@app.route("/home-submit", methods=["POST"])
def home_submit():
    if request.method == "POST":
        try:
            name = request.form["name"]
            if " " not in name:
                raise ValueError("You didn't submit a first and last name")
            lab = request.form["lab"]
            if len(lab) < 3:
                raise ValueError("Your lab name doesn't appear valid")
            email = request.form["email"]
            if len(email) < 3 and "@" not in email:
                raise ValueError("Your email doesn't appear valid")
            now = datetime.now()
            year = now.year
            month = now.month
            day = now.day
            time_24hr = now.strftime("%H:%M:%S")
            with open(os.path.join("data", "Log.txt"), "a") as f:
                f.write(f"{name}, {lab}, {email}, {year}, {month}, {day}, {time_24hr}\n")
            return redirect("/labware-upload")
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            lnum = exc_tb.tb_lineno
            abort(
                400,
                f"{e} \n \n exc_type: {exc_type} \n exc_obj: {exc_obj} \n fname: {fname} \n line_number: {lnum} \n \n traceback: {traceback.format_exc()}",
            )


@app.route("/labware-upload")
def upload_labware_excel():
    return render_template("labware_upload.html")


@app.route("/labware-uploader", methods=["POST"])
def upload_labware_file():
    if request.method == "POST":
        try:
            if request.files["experiment_file"].getbuffer().nbytes > 0:
                file_name = request.files["experiment_file"].filename.split(".")[0]
                df = pd.read_excel(request.files["experiment_file"], skiprows=1, comment='#')
                df = df.dropna(axis=1, how='all')
                df = df.dropna()
            else:
                return "No File submitted", 404

            formatted_time = datetime.now().strftime("%Y%m%d_%H%M%S")

            possible_instruments = ['output_mantis', 'output_integra', 'output_tecan', 'output_hudson']
            
            num_instruments = 0
            for instrument in possible_instruments:
                if instrument in request.form:
                    num_instruments += 1

            
            with tempfile.TemporaryDirectory() as temp_dir:
                if 'output_mantis' in request.form:
                    if num_instruments == 1:
                        output_folder = temp_dir
                    else:
                        os.makedirs(os.path.join(temp_dir, 'Mantis'))
                        output_folder = os.path.join(temp_dir, 'Mantis')

                    ml.mantis_file_maker(df, output_folder)

                if 'output_integra' in request.form:
                    if num_instruments == 1:
                        output_folder = temp_dir
                    else:
                        os.makedirs(os.path.join(temp_dir, 'Integra'))
                        output_folder = os.path.join(temp_dir, 'Integra')
                    il.integra_file_maker(df, output_folder)

                if 'output_tecan' in request.form:
                    if num_instruments == 1:
                        output_folder = temp_dir
                    else:
                        os.makedirs(os.path.join(temp_dir, 'Tecan'))
                        output_folder = os.path.join(temp_dir, 'Tecan')
                    tl.generate_tecan_files(df, output_folder) 

                with tempfile.TemporaryDirectory() as zip_temp_dir:
                    # Zip the processed files
                    output_zip_path = os.path.join(zip_temp_dir, f"processed_files.zip")
                    with zipfile.ZipFile(output_zip_path, "w") as zipf:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                file_path = os.path.join(root, file)
                                zipf.write(
                                    file_path,
                                    os.path.relpath(file_path, temp_dir),
                                )
                    # send new file back
                    with open(output_zip_path, "rb") as f:
                        resp = make_response(f.read())
                        rep_filename = f'processed_labware_files_{formatted_time}_{file_name}.zip'
                        resp.headers["Content-Type"] = "text/plain;charset=UTF-8"
                        resp.headers["Content-Disposition"] = (
                            f"attachment;filename={rep_filename}"
                        )
                        return resp

            

           

        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            lnum = exc_tb.tb_lineno
            abort(
                400,
                f"{e} \n \n exc_type: {exc_type} \n exc_obj: {exc_obj} \n fname: {fname} \n line_number: {lnum} \n \n traceback: {traceback.format_exc()}",
            )
