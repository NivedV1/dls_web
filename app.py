import os, re, shutil
import pandas as pd
from flask import Flask, request, render_template, send_from_directory

UPLOAD = "uploads"
OUTPUT = "output"

os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(OUTPUT, exist_ok=True)

app = Flask(__name__)

def num_sort(name):
    return [int(x) if x.isdigit() else x for x in re.findall(r"\d+|\D+", name)]

def parse_dat(path):
    data = {}

    with open(path) as f:
        txt = f.read()

    try:
        data["Scattering Angle"] = float(re.search(r"Scattering angle:\s*([\d.]+)", txt).group(1))
        data["Mean Size (nm)"] = float(re.search(r"Cumulant 2nd\s+([\d.]+)", txt).group(1))
        data["Count Rate A (kHz)"] = float(re.search(r"Count rate\s*A.*?([\d.]+)", txt).group(1))
        return data
    except:
        return None

@app.route("/", methods=["GET","POST"])
def index():
    files_out = []

    if request.method == "POST":
        shutil.rmtree(UPLOAD, ignore_errors=True)
        os.makedirs(UPLOAD, exist_ok=True)

        for f in request.files.getlist("folder"):
            path = os.path.join(UPLOAD, f.filename)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            f.save(path)

        for sample in os.listdir(UPLOAD):
            sample_path = os.path.join(UPLOAD, sample)
            if not os.path.isdir(sample_path): continue

            for temp in os.listdir(sample_path):
                temp_path = os.path.join(sample_path, temp)

                sheets = {"1mm":[], "5mm":[], "10mm":[], "50mm":[]}

                for file in os.listdir(temp_path):
                    if not file.endswith(".dat"): continue
                    size = re.match(r"(1mm|5mm|10mm|50mm)", file)
                    if not size: continue

                    data = parse_dat(os.path.join(temp_path,file))
                    if not data: continue

                    data["File"] = file
                    sheets[size.group(1)].append(data)

                if not any(sheets.values()):
                    print("Skipping", temp)
                    continue

                out = os.path.join(OUTPUT, f"{sample}_{temp}.xlsx")

                with pd.ExcelWriter(out, engine="openpyxl") as writer:
                    for k,v in sheets.items():
                        if not v: continue
                        v = sorted(v, key=lambda x: num_sort(x["File"]))
                        pd.DataFrame(v).to_excel(writer, sheet_name=k, index=False)

                files_out.append(os.path.basename(out))

    return render_template("index.html", files=files_out)

@app.route("/download/<f>")
def download(f):
    return send_from_directory(OUTPUT, f, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
