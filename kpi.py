import pandas as pd
import datetime as dt
import warnings

warnings.filterwarnings('ignore')

def leer_archivos():
    
    #Nombre y ruta del csv de glpi (los nombres de las variables y de los archivos ahorita son temporales)
    nombre_csv_cris = "cris.csv"
    nombre_csv_pau = "pau.csv"
    nombre_aclaraciones = "Tipos de aclaraciones.xlsx"
    nombre_manuales = "KPI Manuales.xlsx"
    
    #Leer el archivo csv de glpi y el archivo de tipos de aclaraciones
    print("Leyendo archivos")
    df_aclar = pd.read_excel(nombre_aclaraciones)
    df_glpi = pd.read_csv(nombre_csv_cris)
    df_profeco = pd.read_csv(nombre_csv_pau)
    df_manuales = pd.read_excel(nombre_manuales)
    print("Archivos leidos")

    return df_aclar, df_glpi, df_profeco, df_manuales


def eliminar_filas(df_aclar, df_glpi):
    #Se eliminan las que no esten cerradas y las que no son del mes y las que no son de la subcategoria que necesitamos
    
    #Cruce para traerme que tipo de aclaracion es cada una y eliminar los NA
    df_glpi = pd.merge(df_glpi, df_aclar, on="Subcategoria", how='left')
    df_glpi = df_glpi[df_glpi["Tipo de aclaracion"].notna()]

    #Filtrar los que esten cerrados y sean del mes
    df_glpi = df_glpi[df_glpi["Estatus"] == "Cerrado"]
    df_glpi["Fecha de cierre"] = pd.to_datetime(df_glpi["Fecha de cierre"], errors="coerce")
    df_glpi = df_glpi[df_glpi["Fecha de cierre"].dt.month == 1]

    return df_glpi



def tabla_sla(df_glpi, columna):

    table_sla = df_glpi.pivot_table(
        index=columna,
        columns="Dentro de SLA",
        values="Id",
        aggfunc="count"
    )
    
    table_sla = table_sla.div(table_sla.sum(axis=1), axis=0) * 100

    return table_sla



def kpi_servicio(df_glpi):

    #Filtar las que son de servicio y calcular el promedio de dias transcurridos
    df_dias = df_glpi[df_glpi["Tipo de aclaracion"] == "Aclaraciones Servicio"]
    dias_servicio = df_dias["Días transcurridos"].mean()

    return dias_servicio
    

def kpi_procedentes(df_glpi, columna, tipo, calificacion):

    df_glpi = df_glpi[df_glpi[columna] == tipo]
    total = len(df_glpi)
    df_procede = df_glpi[df_glpi["Calificación"] == calificacion]

    
    procede = len(df_procede)
    
    porcentaje = (procede/total)* 100 

    return porcentaje

def kpi_sla(table_sla):

    sla_servicio = table_sla.loc["Aclaraciones Servicio"]["SI"]
    sla_tnr = table_sla.loc["TNR"]["SI"]
    sla_simples_cnr = (table_sla.loc["Aclaraciones Simples"]["SI"] + table_sla.loc["Aclaraciones CNR"]["SI"])/2
    
    return sla_servicio, sla_tnr, sla_simples_cnr

def kpi_sla_sub(table_sla_sub):

    sla_reembolsos = table_sla_sub.loc["Reembolso"]["SI"]
    #Ahorita la subcategoria transpaso por correcion no existe pero el proximo mes si
    #sla_corr_cuenta = table_sla_sub.loc["Traspaso por Correción de cuenta"]["SI"]
    
    return sla_reembolsos#, sla_corr_cuenta

def kpi_sla_total(df_glpi):
    
    total = len(df_glpi)
    df_sla_si = df_glpi[df_glpi["Dentro de SLA"] == "SI"]
    porcentaje = (len(df_sla_si)/total)*100

    return porcentaje

def kpi_no_procede_sub(df_glpi):

    
    df_sub = df_glpi[df_glpi["Subcategoria"] != "Robo/Extravío"]
    total = len(df_sub)
    df_no_procede = df_sub[df_sub["Calificación"] == "No procede"]
    
    porcentaje = (len(df_no_procede)/total)*100

    return porcentaje


def kpi_cerrada_seg_aud(df_profeco):

    #Se verifica si algunas de las columnas de audiencia son de enero
    columnas = ["Audiencia 1", "Audiencia 2", "Audiencia 3", "Audiencia 4"]
    #Se transforman las columnas a fechas
    df_profeco[columnas] = df_profeco[columnas].apply(
        pd.to_datetime, errors="coerce"
    )
    df_profeco = df_profeco[
        (df_profeco[columnas].apply(lambda x: (x.dt.month == 1) & (x.dt.year == 2026)))
        .any(axis=1)
    ]
    
    #Se eliminan las que no esten cerradas en estatus
    df_profeco = df_profeco[df_profeco["Estatus"] == "Cerrado"]

    with pd.ExcelWriter("AAAAA.xlsx",engine="openpyxl") as writer:
        df_profeco.to_excel(writer, "KPI", index=False)

    #Se guarda la cantidad de casos para usarlo mas adelante en el porcentaje
    total = len(df_profeco)

    #Se mantienen los que en la columna audiencia 3 este vacia
    df_profeco = df_profeco[df_profeco["Audiencia 3"].isna()]
    
    #El total de aclaraciones cerradas en segunda audiencia entre el total de aclaraciones cerradas para saber el %
    porcentaje = (len(df_profeco)/total)*100

    return porcentaje
    
    
def calcular_kpi_glpi(df_glpi, table_sla, table_sla_sub, df_profeco, df_manuales):

    kpi = []
    
    #KPI's (se llama a la funcion que calcula cada kpi y se guarda en una variable, al final se guarda nombre:kpi)
    dias_servicio = kpi_servicio(df_glpi)
    
    procede_servicio = kpi_procedentes(df_glpi, "Tipo de aclaracion", "Aclaraciones Servicio", "Procede")
    procede_cnr = kpi_procedentes(df_glpi, "Tipo de aclaracion", "Aclaraciones CNR", "Procede")
    no_procede_reemb = kpi_procedentes(df_glpi, "Subcategoria", "Reembolso", "No procede")
    no_procede_sub = kpi_no_procede_sub(df_glpi) 
    
    sla_servicio, sla_tnr, sla_simples_cnr = kpi_sla(table_sla)
    sla_reembolsos = kpi_sla_sub(table_sla_sub)
    sla_total = kpi_sla_total(df_glpi)

    cerrada_seg_aud = kpi_cerrada_seg_aud(df_profeco)

    #KPI's manuales
    df_manuales["KPI"] = (df_manuales["KPI"]*100).astype(int)
    df_manuales["KPI"] = (df_manuales["KPI"].astype(str)) + "%"
    
    #Se guardan los KPI con sus nombres

    kpi.append(["ESPINOZA VAZQUEZ JIBRAN RAUL", "SLA cierre de tickets general Aclaraciones", str(round(sla_total)) + "%"]) 
    kpi.append(["ESPINOZA VAZQUEZ JIBRAN RAUL", "Improcedencia General aclaraciones excluyendo ROEX", str(round(no_procede_sub)) + "%"]) 
    kpi.append(["VILLEGAS HERRERA CRISTIAN DAVID", "% de tickets procedente Categoría Servicios", str(round(procede_servicio)) + "%"]) 
    kpi.append(["VILLEGAS HERRERA CRISTIAN DAVID", "Días promedio servicios", str(round(dias_servicio))]) 
    kpi.append(["VILLEGAS HERRERA CRISTIAN DAVID", "SLA de cierre de ticket de servicios", str(round(sla_servicio)) + "%"]) 
    kpi.append(["DE LEON CORTES PAULINA", "% de tickets procedente Categoría CNR", str(round(procede_cnr)) + "%"]) 
    kpi.append(["DE LEON CORTES PAULINA", "SLA cierre de tickets A. Simples / CNR", str(round(sla_simples_cnr)) + "%"]) 
    kpi.append(["ZAVALA PONCE ANA PAULA", "% cerradas en segunda audiencia", str(round(cerrada_seg_aud)) + "%"]) 
    kpi.append(["ZAVALA PONCE ANA PAULA", "SLA Aclaraciones TNR", str(round(sla_tnr)) + "%"]) 
    #kpi.append(["ZAVALA PONCE ANA PAULA", "SLA Traspaso corrección de cuenta", str(round(sla_corr_cuenta)) + "%"]) 
    kpi.append(["TEPACH QUINO MIGUEL ANGEL", "% improcedencia reembolsos", str(round(no_procede_reemb)) + "%"]) 
    kpi.append(["TEPACH QUINO MIGUEL ANGEL", "SLA Cierre tickets de reembolsos",str(round(sla_reembolsos)) + "%"]) 

    df_kpi = pd.DataFrame(kpi, columns=["Responsable", "Indicador","KPI"])

    #Se agregan los kpi's manuales a los demas
    df_kpi = pd.concat([df_manuales, df_kpi])

    with pd.ExcelWriter("KPI.xlsx",engine="openpyxl") as writer:
        df_kpi.to_excel(writer, "KPI", index=False)

    print("Archivo exportado")
    
    
def main():
    
    df_aclar, df_glpi, df_profeco, df_manuales = leer_archivos()
    df_glpi = eliminar_filas(df_aclar, df_glpi)

    #Tablas pivotes para el sla
    table_sla = tabla_sla(df_glpi, "Tipo de aclaracion")
    table_sla_sub = tabla_sla(df_glpi, "Subcategoria")

    #KPI's
    calcular_kpi_glpi(df_glpi, table_sla, table_sla_sub, df_profeco, df_manuales)

    
    

main()
