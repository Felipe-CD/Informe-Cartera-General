import pandas as pd
from numpy import nan
from openpyxl import load_workbook

def clean_data(sap, limites, cupos, filtros):
    """
    Función para elinar columnas, realizar filtros, etc, para poder trabahar
    con la base.
    """
    #Eliminación de las columnas con 00 dias de mora
    lista = [x for x in sap.columns.to_list() if 'Cartera A 000' in x]
    for i in lista:
        sap.drop(i, axis=1, inplace=True)
    #Quitar partidas que tengan los filtros digitados por el cliente
    sap["Descripción cabecera pedido"] = sap["Descripción cabecera pedido"].str.lower()
    sap["Descripción cabecera pedido"] = sap["Descripción cabecera pedido"].fillna("NA")
    sap2 = []
    for i in filtros["data"]["filt1"]:
        sap2.append(sap[sap["Descripción cabecera pedido"].str.contains(i)])
        sap.drop(sap[sap["Descripción cabecera pedido"].str.contains(i)].index, inplace=True)
    for i in filtros["data"]["filt2"]:
        sap2.append(sap[(sap["No. Referencia"].notnull()) & (sap["No. Referencia"].str.contains(i))])
        sap.drop(sap[(sap["No. Referencia"].notnull()) & (sap["No. Referencia"].str.contains(i))].index, inplace=True)
    sap.reset_index(drop=True, inplace=True)
    #Las partidas quitadas con los filtros se colocan en un dataframe diferente para dejarlas en una hoja aparte
    sap2 = pd.concat([sap2[i] for i in range(len(sap2))], axis=0)
    sap2.reset_index(drop=True, inplace=True)
    #Buscar acuerdos en el archivo limites para cada partida
    sap["Status"] = sap["No. Identificación Fiscal"].map(limites.drop_duplicates("Nit").set_index("Nit")["ACUERDO"])
    #Colocar "D" en cada partida que tiene la palabra "Cuota" en la columna descripcion cabecera pedido
    sap.loc[(sap["Descripción cabecera pedido"].str.contains("cuota")) & (sap["Ind. Cta Esp."].isnull()), ("Ind. Cta Esp.")] = "D"
    #identificar cada agente que tenga por lo menos una D, y luego a todas las partidas de esos agentes colocarle “ACUERDO” en status
    agentes_acuerdo = sap.loc[sap["Ind. Cta Esp."] == "D",("No. Identificación Fiscal")].drop_duplicates().to_list()
    for i in agentes_acuerdo:
        sap.loc[sap["No. Identificación Fiscal"] == i,("Status")] = "ACUERDO"
    #Reemolazar los vacios en "Status" por "Abierto"
    sap.loc[sap["Status"].isnull(), ("Status")] = "ABIERTO"
    #Si hay una "D" en la columna "Ind. Cta Esp." entonces se coloca un 20 en la columna "Producto"
    #y se coloca "Acuerdo" en la columna "Tipo_producto"
    sap.loc[sap["Ind. Cta Esp."] == "D", ("Producto")] = 20
    sap.loc[sap["Ind. Cta Esp."] == "D", ("Tipo_Producto")] = "Acuerdo"
    #Lo que no sea 20 o 10 en la columna producto entonces debe ser 18
    sap.loc[(sap["Producto"] != 20.0) & (sap["Producto"] != 10.0), ("Producto")] = 18
    #Filtro por dias de mora mayor a 1500 y LATCOM en la columna Descripción
    #incluir latcom y parar si la suma no da 525.659.031
    #Colocar columna "Cartera total" en "Cartera no Vencido"
    #Reemplazar por 0 las columnas "    MAyor a, Cartera Vencida" y por -1 la columna "Días Mora"
    test_latcom = sap.loc[(sap["Descripción"].str.contains("LATCOM")) & (sap["Días Mora"] > 1500), ("         Mayor a")].sum()
    if test_latcom != 525659031:
        check = 'La suma de "   Mayor a" con el filtro (Días mora > 1500 y LATCOM) no es igual a $525.659.031'
    else:
        check = 'Filtro de "   MAyor a" verificado exitosamente'
    sap.loc[(sap["Descripción"].str.contains("LATCOM")) & (sap["Días Mora"] > 1500), ("Cartera No Vencido")] = sap.loc[(sap["Descripción"].str.contains("LATCOM")) & (sap["Días Mora"] > 1500), ("     Cartera Total")]
    sap.loc[(sap["Descripción"].str.contains("LATCOM")) & (sap["Días Mora"] > 1500), ("         Mayor a")] = 0
    sap.loc[(sap["Descripción"].str.contains("LATCOM")) & (sap["Días Mora"] > 1500), (" Cartera Vencida")] = 0
    sap.loc[(sap["Descripción"].str.contains("LATCOM")) & (sap["Días Mora"] > 1500), ("Días Mora")] = -1
    #Buscar region en el archivo limites para cada partida y anexarla despues de columna 2 "No. Identificación Fiscal"
    sap.insert(3, "Region", sap["No. Identificación Fiscal"].map(limites.drop_duplicates("Nit").set_index("Nit")["Nueva Region"]))
    l1 = sap.loc[(sap["Status"] == "ACUERDO") & (sap["Ind. Cta Esp."].isnull()), ("Descripción")].drop_duplicates().to_list()
    l2 = sap.loc[(sap["Status"] == "ACUERDO") & (sap["Ind. Cta Esp."] == "D"), ("Descripción")].drop_duplicates().to_list()
    l3 = [x for x in l1 if x not in l2] #Distribuidores que se cambio de ACUERDO A ABIERTO
    for i in l3:
        sap.loc[sap["Descripción"] == i, ("Status")] = "ABIERTO"

    return sap, limites, cupos, sap2, check, l3

def informe_mexico_120(sap, sap2, l3):
    """
    Funcion para generar las diferentes tablas/estadisticas de los informes
    sap: Base filtrada y lista para trabajar
    sap2: Data extraida para que sap sea la base filtrada/depurada/limpia
    l3: Nombre de los distribuidores a los cuales se le cambio "ACUERDO" por "ABIERTO" (Ya estan al día)
    """
    #CREACION DE TABLAS DINAMICAS
    #MEXICO
    columns_to_analize1 = ['Cartera A 005 Días', 'Cartera A 010 Días', 'Cartera A 015 Días', 'Cartera A 020 Días', 
        'Cartera A 025 Días', 'Cartera A 030 Días']
    columns_to_analize2 = ['Cartera No Vencido', 'Cartera A 060 Días', 'Cartera A 090 Días', 'Cartera A 120 Días', 
            '         Mayor a']
    columns_to_analize11 = ['Cartera A 005 Días', '#5', 'Cartera A 010 Días', '#10', 'Cartera A 015 Días', '#15',
            'Cartera A 020 Días', '#20', 'Cartera A 025 Días', '#25', 'Cartera A 030 Días', '#30']
    sap[columns_to_analize1] = sap[columns_to_analize1].replace({0:nan})
    sap[columns_to_analize2] = sap[columns_to_analize2].replace({0:nan})
    regiones = sap["Region"].drop_duplicates().to_list()
    regiones.sort()
    mexico = pd.DataFrame(columns=["Region"]+columns_to_analize11+columns_to_analize2)
    cont = 0
    for i in regiones:
        total1 = (sap.loc[sap["Region"] == i, (columns_to_analize1)].sum()/1000).round(0)
        cuenta1 = sap.loc[sap["Region"] == i, (['Descripción'] + columns_to_analize1)].groupby("Descripción").sum().replace({0:nan}).count()
        total2 = (sap.loc[sap["Region"] == i, (columns_to_analize2)].sum()/1000).round(0)
        row = [i]
        for j in range(len(total1)):
            row.append(total1[j])
            row.append(cuenta1[j])
        row = row + total2.to_list()
        mexico.loc[cont] = row
        cont += 1
    mid = mexico["Cartera No Vencido"] #colocar columna cartera no vencido al principio
    mexico.drop("Cartera No Vencido", axis=1, inplace=True)
    mexico.insert(1, "ACTUAL", mid)
    mexico.insert(14, "TOT 30", mexico["Cartera A 005 Días"] + mexico["Cartera A 010 Días"] + mexico["Cartera A 015 Días"] + mexico["Cartera A 020 Días"] + mexico["Cartera A 025 Días"] + mexico["Cartera A 030 Días"]) #insertar columna de la suma de cartera al final de 30 dias

    #CARTERA 120 DIAS
    cartera_120 = sap.pivot_table(index="Descripción", values=["         Mayor a"], aggfunc=["sum"], columns=["Region"], margins=True)
    cartera_120 = cartera_120[cartera_120[('sum', '         Mayor a', 'All')] > 0]

    #Exportar sap, sap2, mexico y 120 dias a un archivo de excel en hojas diferentes
    writer = pd.ExcelWriter("Informe_1.xlsx", engine="xlsxwriter")
    workbook = writer.book
    worksheet = workbook.add_worksheet("Base depurada")
    writer.sheets["Base depurada"] = worksheet
    worksheet.write_string(0, 0, f"Base depurada con los filtros especificados en el programa: {filtros['data']['filt1']}")
    worksheet.write_string(1, 0, f"Los distribuidores a los cuales se les cambio el status de 'ACUERDO' a 'ABIERTO' son: {', '.join(str(x) for x in l3)}")
    sap.to_excel(writer, sheet_name="Base depurada", startrow=3 , startcol=1, index=False)
    worksheet = workbook.add_worksheet("Data extraida")
    writer.sheets["Data extraida"] = worksheet
    worksheet.write_string(0, 0, "DATA EXTRAIDA")
    sap2.to_excel(writer, sheet_name="Data extraida", startrow=2 , startcol=1, index=False)
    worksheet = workbook.add_worksheet("Informe Mexico")
    writer.sheets["Informe Mexico"] = worksheet
    worksheet.write_string(0, 1, "Comunicación Celular Comcel S.A")
    worksheet.write_string(1, 1, "INTEGRACIÓN DE LA CUENTA POR COBRAR A")
    worksheet.write_string(2, 1, "Canales de distribución regiones x-x")
    worksheet.write_string(3, 1, "CIFRAS EN MILES DE PESOS")
    mexico.to_excel(writer, sheet_name="Informe Mexico", startrow=10, startcol=1, index=False)
    worksheet = workbook.add_worksheet("Cartera 120 dias")
    writer.sheets["Cartera 120 dias"] = worksheet
    worksheet.write_string(0, 0, "Informe cartera 120 dias")
    cartera_120.to_excel(writer, sheet_name="Cartera 120 dias", startrow=2, startcol=1)
    writer.save()
    return

def cartera_general(sap,sap2, cupos, cerrados, cerrados_sap):
    """
    Esta función genera las tablas y tabl;as dinamicas que se abren del archivo 
    cartera general y las actualiza
    """
    def generate_pivot_table(data, filtro):
        """
        Función que retorna una tabla dinamica con un filtro en la columna Status y columnas desde Cartera no vencida hasta Mayor a, incluyendo suma desde a 10 hasta >120


        data: Dataframe como la base filtrada con filtros aplicados de Producto, region, etc

        filtro: "ACUERDO" o "ABIERTO"
        """
        df = data[(data["Status"] == filtro)].pivot_table(
                        index=["No. de Cliente","No. Identificación Fiscal","Descripción"],
                        values=["Cartera No Vencido","Cartera a 10 Dias","Cartera a 20 Dias","Cartera a 30 Dias","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","         Mayor a","     Cartera Total"], 
                        aggfunc=["sum"])
        df.columns = [j for i,j in df.columns]
        df = df[["Cartera No Vencido","Cartera a 10 Dias","Cartera a 20 Dias","Cartera a 30 Dias","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","         Mayor a","     Cartera Total"]] #Reordenar el orden de las columnas
        df.insert(8, "Total Vencida", df.iloc[:, 1:8].sum(axis=1))
        df.reset_index(inplace=True)
        return df

    #Agrupo y dejo solo columnas de dias de cartera 10, 20, 30
    sap.insert(17, "Cartera a 10 Dias", sap[["Cartera A 005 Días","Cartera A 010 Días"]].sum(axis=1, min_count=1))
    sap.insert(20, "Cartera a 20 Dias", sap[["Cartera A 015 Días","Cartera A 020 Días"]].sum(axis=1, min_count=1))
    sap.insert(23, "Cartera a 30 Dias", sap[["Cartera A 025 Días","Cartera A 030 Días"]].sum(axis=1, min_count=1))
    sap.drop(["Cartera A 005 Días","Cartera A 010 Días","Cartera A 015 Días","Cartera A 020 Días","Cartera A 025 Días","Cartera A 030 Días"], axis=1, inplace=True)
    if sap["     Cartera Total"].sum() == sap["Cartera No Vencido"].sum() + sap["Cartera a 10 Dias"].sum() + sap["Cartera a 20 Dias"].sum() + sap["Cartera a 30 Dias"].sum() + sap["Cartera A 060 Días"].sum() + sap["Cartera A 090 Días"].sum() + sap["Cartera A 120 Días"].sum() + sap["         Mayor a"].sum():
        check = "La suma desde 'Cartera total' hasta 'mayor a' da igual que 'Cartera Total'"
    else:
        check = "La suma desde 'Cartera total' hasta 'mayor a' NO da igual que 'Cartera Total'"
    
    #CREACION DE TABLAS Y VARIABLES PARA MODIFICAR EL INFORME
    #sap2 equivale a "Data_Conceptos_Excluyentes"
    #par_exclu_inte equivale a la data en "Partidas Excluidas intereses"
    par_exclu_inte = sap2.loc[
                            (sap2["Descripción cabecera pedido"].str.contains("valor presente neto")) | 
                            (sap2["Descripción cabecera pedido"].str.contains("vpn")),
                            ("No. de Cliente","Descripción","No. Identificación Fiscal","Cartera No Vencido","Días Mora"," Cartera Vencida","     Cartera Total")
    ]
    par_exclu_inte["Descripción_2"] = "COMPENSACION INTERESES PRESTAMO DIS"
    #suma_recargas equivale a la celda (Conceptos Recargas en Línea) de la hoja "Partidas Excluidas intereses"
    suma_recargas = sap2.loc[
                            (sap2["Descripción cabecera pedido"].str.contains("recarga en linea")) |
                            (sap2["Descripción cabecera pedido"].str.contains("recarga $")) |
                            (sap2["Descripción cabecera pedido"].str.contains("recarga  $")),
                            ("     Cartera Total")
    ].sum()
    #R1, R2, R3, R4, R5 para colocar ne las respectivas hojas del libro de excel
    regiones = sap["Region"].drop_duplicates().to_list()
    regiones.sort()
    R1 = sap[sap["Region"] == regiones[0]]
    R2 = sap[sap["Region"] == regiones[1]]
    R3 = sap[sap["Region"] == regiones[2]]
    R4 = sap[sap["Region"] == regiones[3]]
    R5 = sap[sap["Region"] == regiones[4]]
    #Data para las paginas "detalle otro concepto abierto" y "detalle kits"
    otro_concepto_abierto = sap[sap["Producto"] == 18]
    detalle_kits = sap[sap["Producto"] == 10]
    #TABLAS DINAMICAS
    #Hoja Informe_Acuerdos
    informe_acuerdos = sap[(sap["Status"] == "ACUERDO") & (sap["Ind. Cta Esp."].isnull())].pivot_table(
                        index=["No. de Cliente","No. Identificación Fiscal","Descripción"],
                        values=["Cartera No Vencido","Cartera a 10 Dias","Cartera a 20 Dias","Cartera a 30 Dias","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","         Mayor a","     Cartera Total"], 
                        aggfunc=["sum"]) #Esta dataframe se debe imprimir con una fila antes, debido a que al quitar el header se imprime una despues
    informe_acuerdos.columns = [j for i,j in informe_acuerdos.columns]
    informe_acuerdos = informe_acuerdos[["Cartera No Vencido","Cartera a 10 Dias","Cartera a 20 Dias","Cartera a 30 Dias","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","         Mayor a","     Cartera Total"]] #Reordenar el orden de las columnas
    info_temporal = sap[(sap["Status"] == "ACUERDO") & (sap["Ind. Cta Esp."] == "D")].pivot_table(
                        index=["No. de Cliente","No. Identificación Fiscal","Descripción"],
                        values=["     Cartera Total"], 
                        aggfunc=["sum"]) #Esta dataframe se debe imprimir con una fila antes, debido a que al quitar el header se imprime una despues
    informe_acuerdos.insert(8, "Total Vencida", informe_acuerdos.iloc[:, 1:8].sum(axis=1))
    informe_acuerdos["Acuerdo"] = info_temporal[("sum","     Cartera Total")]
    informe_acuerdos.reset_index(inplace=True)
    #HOJA "Kit abiertos"
    #primera tabla "abierto"
    kit_abiertos = detalle_kits[(detalle_kits["Status"] == "ABIERTO")].pivot_table(
                        index=["No. de Cliente","No. Identificación Fiscal","Descripción"],
                        values=["Cartera No Vencido","Cartera a 10 Dias","Cartera a 20 Dias","Cartera a 30 Dias","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","         Mayor a","     Cartera Total"], 
                        aggfunc=["sum"])
    kit_abiertos.columns = [j for i,j in kit_abiertos.columns]
    kit_abiertos = kit_abiertos[["Cartera No Vencido","Cartera a 10 Dias","Cartera a 20 Dias","Cartera a 30 Dias","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","         Mayor a","     Cartera Total"]] #Reordenar el orden de las columnas
    kit_abiertos.insert(8, "Total Vencida", kit_abiertos.iloc[:, 1:8].sum(axis=1))
    kit_abiertos.reset_index(inplace=True) #Pasar los multiindex a columnas con su respectivo nombre
    kit_abiertos.insert(3, "Limite de credito", kit_abiertos["No. de Cliente"].map(cupos.drop_duplicates("Cliente").set_index("Cliente")["Límite crédito"]))
    kit_abiertos.insert(4, "Extra cupo", 0)
    #segunda tabla "acuerdo"
    kit_acuerdo = detalle_kits[(detalle_kits["Status"] == "ACUERDO")].pivot_table(
                        index=["No. de Cliente","No. Identificación Fiscal","Descripción"],
                        values=["Cartera No Vencido","Cartera a 10 Dias","Cartera a 20 Dias","Cartera a 30 Dias","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","         Mayor a","     Cartera Total"], 
                        aggfunc=["sum"])
    kit_acuerdo.columns = [j for i,j in kit_acuerdo.columns]
    kit_acuerdo = kit_acuerdo[["Cartera No Vencido","Cartera a 10 Dias","Cartera a 20 Dias","Cartera a 30 Dias","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","         Mayor a","     Cartera Total"]] #Reordenar el orden de las columnas
    kit_acuerdo.insert(8, "Total Vencida", kit_acuerdo.iloc[:, 1:8].sum(axis=1))
    kit_acuerdo.reset_index(inplace=True) #Pasar los multiindex a columnas con su respectivo nombre
    kit_acuerdo.insert(3, "Limite de credito", kit_acuerdo["No. de Cliente"].map(cupos.drop_duplicates("Cliente").set_index("Cliente")["Límite crédito"]))
    kit_acuerdo.insert(4, "Extra cupo", 0)
    #HOJA "Otros conceptos abiertos"
    otros_conceptos_abiertos = generate_pivot_table(otro_concepto_abierto, "ABIERTO")
    otros_conceptos_acuerdos = generate_pivot_table(otro_concepto_abierto, "ACUERDO")
    #HOJA "Agentes R4 Centro-Oriente"
    r4_abiertos = generate_pivot_table(R4, "ABIERTO")
    r4_acuerdo = generate_pivot_table(R4, "ACUERDO")
    #HOJA "Agentes R3"
    r3_abiertos = generate_pivot_table(R3, "ABIERTO")
    r3_acuerdo = generate_pivot_table(R3, "ACUERDO")
    #HOJA "Agentes R2"
    r2_abiertos = generate_pivot_table(R2, "ABIERTO")
    r2_acuerdo = generate_pivot_table(R2, "ACUERDO")
    #HOJA "Agentes R1"
    r1_abiertos = generate_pivot_table(R1, "ABIERTO")
    r1_acuerdo = generate_pivot_table(R1, "ACUERDO")

    #CERRADOS
    #Data cerrados
    cerrados_sap["Zona"] = cerrados_sap["No. Identificación Fiscal"].map(cerrados.drop_duplicates("NIT").set_index("NIT")["zona"])
    #Agentes cerrados (imprimir uno bajo el otro)
    cerrados_table_co03 = cerrados_sap[cerrados_sap["Zona"] == "CO03"].pivot_table(
                            index=["No. de Cliente","Descripción"],
                            values=["Cartera No Vencido","Cartera A 010 Días","Cartera A 020 Días","Cartera A 030 Días","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","          Mayor a","    Cartera Total"],
                            aggfunc=["sum"])
    cerrados_table_co03.columns = [j for i,j in cerrados_table_co03.columns]
    cerrados_table_co03 = cerrados_table_co03[["Cartera No Vencido","Cartera A 010 Días","Cartera A 020 Días","Cartera A 030 Días","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","          Mayor a","    Cartera Total"]]
    cerrados_table_co03.insert(8, "Total Vencida", cerrados_table_co03.iloc[:, 1:8].sum(axis=1))
    cerrados_table_co03.reset_index(inplace=True)

    cerrados_table_co04 = cerrados_sap[cerrados_sap["Zona"] == "CO04"].pivot_table(
                            index=["No. de Cliente","Descripción"],
                            values=["Cartera No Vencido","Cartera A 010 Días","Cartera A 020 Días","Cartera A 030 Días","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","          Mayor a","    Cartera Total"],
                            aggfunc=["sum"])
    cerrados_table_co04.columns = [j for i,j in cerrados_table_co04.columns]
    cerrados_table_co04 = cerrados_table_co04[["Cartera No Vencido","Cartera A 010 Días","Cartera A 020 Días","Cartera A 030 Días","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","          Mayor a","    Cartera Total"]]
    cerrados_table_co04.insert(8, "Total Vencida", cerrados_table_co04.iloc[:, 1:8].sum(axis=1))
    cerrados_table_co04.reset_index(inplace=True)

    cerrados_table_co05 = cerrados_sap[cerrados_sap["Zona"] == "CO05"].pivot_table(
                            index=["No. de Cliente","Descripción"],
                            values=["Cartera No Vencido","Cartera A 010 Días","Cartera A 020 Días","Cartera A 030 Días","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","          Mayor a","    Cartera Total"],
                            aggfunc=["sum"])
    cerrados_table_co05.columns = [j for i,j in cerrados_table_co05.columns]
    cerrados_table_co05 = cerrados_table_co05[["Cartera No Vencido","Cartera A 010 Días","Cartera A 020 Días","Cartera A 030 Días","Cartera A 060 Días","Cartera A 090 Días","Cartera A 120 Días","          Mayor a","    Cartera Total"]]
    cerrados_table_co05.insert(8, "Total Vencida", cerrados_table_co05.iloc[:, 1:8].sum(axis=1))
    cerrados_table_co05.reset_index(inplace=True)

    #GENERACION DE INFORME EN UN ARCHIVO DE EXCEL -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    writer = pd.ExcelWriter("Informe_Cartera_General_Program.xlsx", engine="xlsxwriter")
    workbook = writer.book
    #"Data_Conceptos_Excluyentes"
    worksheet = workbook.add_worksheet("Data_Conceptos_Excluyentes")
    writer.sheets["Data_Conceptos_Excluyentes"] = worksheet
    sap2.to_excel(writer, sheet_name="Data_Conceptos_Excluyentes", index=False)
    #Partidas Excluidas intereses
    worksheet = workbook.add_worksheet("Partidas Excluidas intereses")
    writer.sheets["Partidas Excluidas intereses"] = worksheet
    worksheet.write_string(1,2, "Compensaciones prestamos")
    par_exclu_inte.to_excel(writer, "Partidas Excluidas intereses", index=0, startrow=4, startcol=2)
    worksheet.write_string(par_exclu_inte.shape[0] + 6, 2, "Conceptos  Recargas en Línea")
    worksheet.write_number(par_exclu_inte.shape[0] + 6, 3, suma_recargas)
    #DATA R1
    worksheet = workbook.add_worksheet("DATA R1")
    writer.sheets["DATA R1"] = worksheet
    R1.to_excel(writer, sheet_name="DATA R1", index=False)
    #DATA R2
    worksheet = workbook.add_worksheet("DATA R2")
    writer.sheets["DATA R2"] = worksheet
    R2.to_excel(writer, sheet_name="DATA R2", index=False)
    #DATA R3
    worksheet = workbook.add_worksheet("DATA R3")
    writer.sheets["DATA R3"] = worksheet
    R3.to_excel(writer, sheet_name="DATA R3", index=False)
    #DATA R4
    worksheet = workbook.add_worksheet("DATA R4")
    writer.sheets["DATA R4"] = worksheet
    R4.to_excel(writer, sheet_name="DATA R4", index=False)
    #DATA R5
    worksheet = workbook.add_worksheet("DATA R5")
    writer.sheets["DATA R5"] = worksheet
    R5.to_excel(writer, sheet_name="DATA R5", index=False)
    #Detalle otro conpectos abiertos
    worksheet = workbook.add_worksheet("Detalle otros conceptos abierto")
    writer.sheets["Detalle otros conceptos abierto"] = worksheet
    otro_concepto_abierto.to_excel(writer, sheet_name="Detalle otros conceptos abierto", index=False)
    #Detalle kits
    worksheet = workbook.add_worksheet("Detalle kits")
    writer.sheets["Detalle kits"] = worksheet
    detalle_kits.to_excel(writer, sheet_name="Detalle kits", index=False)
    #Informe_Acuerdos
    worksheet = workbook.add_worksheet("Informe_Acuerdos")
    writer.sheets["Informe_Acuerdos"] = worksheet
    worksheet.write_string(0, 1, "INFORME CARTERA DE DISTRIBUIDORES EN ACUERDO DE PAGO CON CORTE")
    informe_acuerdos.to_excel(writer, sheet_name="Informe_Acuerdos", startcol=2, startrow=4)
    #Kit abiertos
    worksheet = workbook.add_worksheet("Kits abiertos")
    writer.sheets["Kits abiertos"] = worksheet
    worksheet.write_string(1, 1, "CARTERA DISTRIBUIDORES KIT A CORTE")
    worksheet.write_string(4, 1, "DISTRIBUIDORES ABIERTOS")
    kit_abiertos.to_excel(writer, sheet_name="Kits abiertos", index=False, startrow=5, startcol=0)
    worksheet.write_string(kit_abiertos.shape[0] + 7, 2, "DISTRIBUIDORES EN ACUERDO DE PAGO CON CARTERA ADICIONAL")
    kit_acuerdo.to_excel(writer, sheet_name="Kits abiertos", index=False, startrow=kit_abiertos.shape[0] + 8, startcol=0)
    #Otros conceptos abiertos
    worksheet = workbook.add_worksheet("Otros conceptos abiertos")
    writer.sheets["Otros conceptos abiertos"] = worksheet
    worksheet.write_string(1, 1, "CARTERA  OTROS CONCEPTOS DISTRIBUIDORES ABIERTOS  CON CORTE ")
    worksheet.write_string(4, 1, "DISTRIBUIDORES ABIERTOS")
    otros_conceptos_abiertos.to_excel(writer, sheet_name="Otros conceptos abiertos", index=False, startrow=5, startcol=0)
    worksheet.write_string(otros_conceptos_abiertos.shape[0] + 7, 2, "DISTRIBUIDORES EN ACUERDO DE PAGO CON CARTERA ADICIONAL")
    otros_conceptos_acuerdos.to_excel(writer, sheet_name="Otros conceptos abiertos", index=False, startrow=otros_conceptos_abiertos.shape[0] + 8, startcol=0)
    #Agentes R4 Centro-Oriente
    worksheet = workbook.add_worksheet("Agentes R4 Centro-Oriente")
    writer.sheets["Agentes R4 Centro-Oriente"] = worksheet
    worksheet.write_string(1, 1, "CARTERA DISTRIBUIDORES  ZONA ORIENTE CON CORTE ")
    worksheet.write_string(4, 1, "DISTRIBUIDORES ABIERTOS")
    r4_abiertos.to_excel(writer, sheet_name="Agentes R4 Centro-Oriente", index=False, startrow=5, startcol=0)
    worksheet.write_string(r4_abiertos.shape[0] + 7, 2, "DISTRIBUIDORES EN ACUERDO DE PAGO CON CARTERA ADICIONAL")
    r4_acuerdo.to_excel(writer, sheet_name="Agentes R4 Centro-Oriente", index=False, startrow=r4_abiertos.shape[0] + 8, startcol=0)
    #Agentes R3 SurOccidente
    worksheet = workbook.add_worksheet("Agentes R3 SurOccidente")
    writer.sheets["Agentes R3 SurOccidente"] = worksheet
    worksheet.write_string(1, 1, "CARTERA DISTRIBUIDORES ZONA OCCIDENTE CON CORTE ")
    worksheet.write_string(4, 1, "DISTRIBUIDORES ABIERTOS")
    r3_abiertos.to_excel(writer, sheet_name="Agentes R3 SurOccidente", index=False, startrow=5, startcol=0)
    worksheet.write_string(r3_abiertos.shape[0] + 7, 2, "DISTRIBUIDORES EN ACUERDO DE PAGO CON CARTERA ADICIONAL")
    r3_acuerdo.to_excel(writer, sheet_name="Agentes R3 SurOccidente", index=False, startrow=r3_abiertos.shape[0] + 8, startcol=0)
    #Agentes R2 NorOccidente
    worksheet = workbook.add_worksheet("Agentes R2 NorOccidente")
    writer.sheets["Agentes R2 NorOccidente"] = worksheet
    worksheet.write_string(1, 1, "CARTERA DISTRIBUIDORES ZONA R2 CON CORTE ")
    worksheet.write_string(4, 1, "DISTRIBUIDORES ABIERTOS")
    r2_abiertos.to_excel(writer, sheet_name="Agentes R2 NorOccidente", index=False, startrow=5, startcol=0)
    worksheet.write_string(r2_abiertos.shape[0] + 7, 2, "DISTRIBUIDORES EN ACUERDO DE PAGO CON CARTERA ADICIONAL")
    r2_acuerdo.to_excel(writer, sheet_name="Agentes R2 NorOccidente", index=False, startrow=r2_abiertos.shape[0] + 8, startcol=0)
    #Agentes R1 Costa
    worksheet = workbook.add_worksheet("Agentes R1 Costa")
    writer.sheets["Agentes R1 Costa"] = worksheet
    worksheet.write_string(1, 1, "CARTERA DISTRIBUIDORES ZONA COSTA CON CORTE ")
    worksheet.write_string(4, 1, "DISTRIBUIDORES ABIERTOS")
    r1_abiertos.to_excel(writer, sheet_name="Agentes R1 Costa", index=False, startrow=5, startcol=0)
    worksheet.write_string(r1_abiertos.shape[0] + 7, 2, "DISTRIBUIDORES EN ACUERDO DE PAGO CON CARTERA ADICIONAL")
    r1_acuerdo.to_excel(writer, sheet_name="Agentes R1 Costa", index=False, startrow=r1_abiertos.shape[0] + 8, startcol=0)
    #Data cerrados
    worksheet = workbook.add_worksheet("Data cerrados")
    writer.sheets["Data cerrados"] = worksheet
    cerrados_sap.to_excel(writer, sheet_name="Data cerrados", index=False)
    #Agentes cerrados
    worksheet = workbook.add_worksheet("Agentes cerrados")
    writer.sheets["Agentes cerrados"] = worksheet
    worksheet.write_string(1, 1, "CARTERA DISTRIBUIDORES CERRADOS A CORTE ")
    worksheet.write_string(4, 1, "DISTRIBUIDORES CERRADOS ORIENTE")
    cerrados_table_co03.to_excel(writer, sheet_name="Agentes cerrados", index=False, startrow=5, startcol=0)
    worksheet.write_string(cerrados_table_co03.shape[0] + 7, 2, "DISTRIBUIDORES CERRADOS OCCIDENTE")
    cerrados_table_co04.to_excel(writer, sheet_name="Agentes cerrados", index=False, startrow=cerrados_table_co03.shape[0] + 8, startcol=0)
    worksheet.write_string(cerrados_table_co03.shape[0] + 8 + cerrados_table_co04.shape[0] + 7, 2, "DISTRIBUIDORES CERRADOS COSTA")
    cerrados_table_co05.to_excel(writer, sheet_name="Agentes cerrados", index=False, startrow=cerrados_table_co03.shape[0] + 8 + cerrados_table_co04.shape[0] + 8, startcol=0)
    writer.save()
    return check
