from flask import Flask, render_template, redirect, flash, request, url_for, session
from flaskwebgui import FlaskUI
from functions import clean_data, informe_mexico_120, cartera_general
import pandas as pd
from numpy import nan

def read_files(name_sap, name_limites, name_cupos):
    """
    Esta función lee los archivos seleccionados por el usuario
    en la interfaz grafica y verifica si existe un error en ellos
    """
    #Nombres de columna de los archivos
    sap_headers = ['No. de Cliente', 'Descripción', 'No. Identificación Fiscal', 'Clase Doc.', 'Ind. Cta Esp.',
        'No. Referencia', 'No. Factura', 'No. Doc. Contable', 'Fecha Contabilización', 'Fecha Documento', 
        'Entrada Documento', 'Fecha Base', 'Fecha Vencimiento', 'Cartera No Vencido', 'Cartera A 005 Días', 
        'Cartera A 010 Días', 'Cartera A 015 Días', 'Cartera A 020 Días', 'Cartera A 025 Días', 'Cartera A 030 Días', 
        'Cartera A 060 Días', 'Cartera A 090 Días', 'Cartera A 120 Días', '         Mayor a', 'Días Mora', 
        ' Cartera Vencida', '     Cartera Total', 'Producto', 'Tipo_Producto', '%Participación', 'Zona de Ventas', 
        'Descripción.1', 'Descripción cabecera pedido', 'Organización de Ventas', 'Descripción.2', 
        'Canal de Distribución', 'Descripción.3', 'Clase Pedido', 'Descr.Clase', 'No.Pedido', 'Inactivo', 
        'Cuenta', 'Gr.Cliente', 'Gr.Cliente.1']
    limites_headers = ['Código', 'Nit', 'DISTRIBUIDORES', 'Código.1', 'BLOQUEOS', 'ACUERDO', 'fecha Acuer', 'ZONA', 
        'poliza', 'ZONA.1', 'antigua region', 'Nueva Region', 'ANALISTA COMISIONES', 'JEFE CANAL', 'JEFE TERRITORIO', 
        'GERENTE CANAL', 'lider', 'CATEGORIA', 'para validar']
    cupos_headers = ['Cliente', 'ACC', 'Cta.créd.', 'Cl.riesgo', 'Mon.', 'Límite crédito', ' Comprom.total', 'Agotamiento']
    #Lectura
    flag = False
    sap = pd.read_excel(name_sap, skiprows=[0,1,2,3,4,5,6,7,9], usecols="B:AV")
    cupos = pd.read_excel("cupos prepago.xlsx", skiprows=[0,1,2,4], usecols=[1,3,4,5,6,7,8,9])
    limites = pd.ExcelFile(name_limites)
    if "LIMITES" not in limites.sheet_names:
        flag = True
        flash(f"No existe la hoja (LIMITES) en el archivo {name_limites}")
    else:
        limites = pd.read_excel(name_limites, sheet_name="LIMITES")
    #Verificación de errores
    for i in sap_headers:
        if i not in sap.columns.to_list():
            flag = True
            flash(f"No existe la columna ({i}) en el archivo SAP")
    for i in limites_headers:
        if i not in limites.columns.to_list():
            flag = True
            flash(f"No existe la columna ({i}) en el archivo Limites")
    for i in cupos_headers:
        if i not in cupos.columns.to_list():
            flag = True
            flash(f"No existe la columna ({i}) en el archivo Cupos")

    return sap, limites, cupos, flag


app = Flask(__name__)
app.secret_key = "super_secreto"
#ui = FlaskUI(app, width=1020, height=650)                 # Creacion de la IU (UI)
def shutdown_server():
    func = request.environ.get('werkzeug.server.shutdown')
    if func is None:
        raise RuntimeError('Not running with the Werkzeug Server')
    func()

@app.route("/", methods=["GET", "POST"])
def home():
    """
    Esta pagina es la de inicio, en donde se seleccionan los
    archivos a trabajar
    """
    global sap, limites, cupos, sap2
    if request.method == "POST":
        name_sap = request.files["name_sap"]
        name_limites = request.files["name_limit"]
        name_cupos = request.files["name_cupos"]
        salida = request.form["namesal"] + "xlsx"
        sap, limites, cupos, flag = read_files(name_sap, name_limites, name_cupos) #Lectura de los archivos
        if flag == True:
            return redirect(url_for("error"))
        else:
            return redirect(url_for("filtres"))
    else:
        return render_template("read.html")

@app.route("/error")
def error():
    return render_template("errors.html")

@app.route("/filtres", methods=["GET", "POST"])
def filtres():
    """
    Esta pagina muestra los filtroa a aplicar a la columna
    "Descripción cabecera pedido" del archivo SAP
    """
    global filtros
    filtros = {
        "data":
            {
                "name1": "Descripción cabecera pedido",
                "filt1": "castigo,valor presente neto,recarga en linea,vpn",
                "name2": "No. Referencia",
                "filt2": "CASTIGO"
            }
    }
    if request.method == "POST":
        #Leer los filtros
        filtros["data"]["filt1"] = list(request.form["lista1"].split(","))
        filtros["data"]["filt2"] = list(request.form["lista2"].split(","))
        lista_filtros1 = request.form["lista1"]
        lista_filtros2 = request.form["lista2"]
        session["lista_filtros1"] = lista_filtros1
        session["lista_filtros2"] = lista_filtros2
        return redirect(url_for("execute"))    
    else:
        return render_template("filtres.html", f=filtros)

@app.route("/execute")
def execute():
    global sap, limites, cupos, sap2, filtros, check1, check2
    sap, limites, cupos, sap2, check1 = clean_data(sap, limites, cupos, filtros)
    flash(check1)
    informe_mexico_120(sap, sap2)
    check2 = cartera_general(sap, sap2)
    flash(check2)
    return render_template("execute.html")

@app.route('/shutdown', methods=['GET'])
def shutdown():
    shutdown_server()
    return render_template("close.html")

app.run(debug=True)
#ui.run()
