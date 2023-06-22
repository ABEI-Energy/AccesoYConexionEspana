# from docx import Document  # for accessing the document by python-docx
import datetime as dt
import io
import locale as lc
from zipfile import ZipFile

import pandas as pd
import streamlit as st
import streamlit_toggle as tog
from docx.shared import Cm
import PyPDF2 

import coordinates as cd
import numToLet as ntl
import post as ps
import readers as rd
import writers as wt
from aereas import *


def duplicateDoc(filemodelo):
    from docx import Document  # for accessing the document by python-docx

    global edited_doc1
    # access Word document"MODELOS\Modelo anteproyecto AFL.docx"
    my_file = Document(filemodelo)

    return my_file

#Set the language for datetime
lc.setlocale(lc.LC_ALL,'es_ES.UTF-8')
month = dt.datetime.now().strftime("%B %Y")
def normalize(string):
    return str(round(float(string.replace(",", ".")),2))

def normalize2(string):
    return str(string.replace(",", "."))

st.set_page_config(layout="wide")

def resize_image(image, max_width):
    width, height = image.size
    if width > max_width:
        ratio = max_width / width
        new_width = max_width
        new_height = int(height * ratio)
        return image.resize((new_width, new_height))
    return image


def insert_image_in_cell(doc, dic, picFile):
    if rootLogos not in dic['logoC']:
        dic['logoC'] = rootLogos + "/" + dic['logoC']
    if rootLogos not in dic['logoH']:
        dic['logoH'] = rootLogos + "/" + dic['logoH']
    import io

    from PIL import Image

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "logoC" in cell.text:
                    imagen = Image.open(dic['logoC'])

                    logoC = io.BytesIO()
                    imagen.save(logoC, format = 'PNG')
                    logoC.seek(0)
                    cell.text = ""
                    cell_paragraph = cell.paragraphs[0]
                    run = cell_paragraph.add_run()
                    run.add_picture(logoC, width = Cm(7.0), height = Cm(4.5))

                if ("tomaAerea" in cell.text) and (dic['tomaAerea'] != "pass"):
                    imagen = Image.open(picFile)
                    size = imagen.size
                    imWidth = float(size[0])
                    imHeight = float(size[1])                    
                    if imWidth > 16.55:
                        imHeight = imHeight*16.55/imWidth
                        imWidth = 16.55
                    logoC = io.BytesIO()
                    imagen.save(logoC, format = 'PNG')
                    logoC.seek(0)
                    cell.text = ""
                    cell_paragraph = cell.paragraphs[0]
                    run = cell_paragraph.add_run()
                    run.add_picture(logoC, width = Cm(imWidth), height = Cm(imHeight))

                if ("figuraStruct" in cell.text) and (dic['figuraStruct'] != "pass"):
                    imagen = Image.open(dic['figuraStruct'])
                    logoC = io.BytesIO()
                    imagen.save(logoC, format = 'PNG')
                    logoC.seek(0)
                    cell.text = ""
                    cell_paragraph = cell.paragraphs[0]
                    run = cell_paragraph.add_run()
                    run.add_picture(logoC, width = Cm(6.95), height = Cm(5.1))          

                if ("figuraCronograma" in cell.text) and (dic['figuraCronograma'] != "pass"):
                    imagen = Image.open(dic['figuraCronograma'])
                    logoC = io.BytesIO()
                    imagen.save(logoC, format = 'PNG')
                    logoC.seek(0)
                    cell.text = ""
                    cell_paragraph = cell.paragraphs[0]
                    run = cell_paragraph.add_run()
                    if int(dic['mesesProjDur']) == 10:
                        run.add_picture(logoC, width = Cm(15.2), height = Cm(8.5))   
                    
                    else:
                        run.add_picture(logoC, width = Cm(15.2), height = Cm(14.3))   
                    
    for section in doc.sections:
        for table in section.header.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "logoH" in cell.text:
                        imagen = Image.open(dic['logoC'])
                        logoC = io.BytesIO()
                        imagen.save(logoC, format = 'PNG')
                        logoC.seek(0)
                        cell.text = ""
                        cell_paragraph = cell.paragraphs[0]
                        run = cell_paragraph.add_run()
                        run.add_picture(logoC, width = Cm(2.75), height = Cm(1.75))          

    flagPicWriter = 1
    return flagPicWriter 


"""
# Documentos Acceso y Conexión España
"""

user = st.text_input("Nombre redactor del documento")
if user:
    user = user.strip().split(' ')
    auxUser = []
    for word in user:
        auxUser.append([*word][0])
    user = '.'.join(auxUser)


# Create a radio button with three choices
proyectoTipo = st.radio("Selecciona tipo de instalación del proyecto:", ("Fotovoltáico", "Eólico", "Hidrógeno"))

mainDic = {}

mainDic['nombreProyecto'] = st.text_input("Nombre del proyecto")
# Sociedades
rootSociedades = 'SOCIEDADES/Sociedades España.csv'
rootEstructuras = "DATASHEETS/Estructuras"
rootWord = "MODELOS"
rootLogos = "SOCIEDADES/Logos"
csvTensiones = 'DATASHEETS/Tensiones maximas.csv'
csvTrafos = 'DATASHEETS/Trafos/Datasheet Trafos.csv'
csvModulos = 'DATASHEETS/Módulos/Datasheet Módulos.csv'
csvInverters = 'DATASHEETS/Inversores/Datasheet Inversores.csv'
csvCeldas = 'DATASHEETS/CeldasMT/CeldasMT.csv'


rootSampleFiles = "SAMPLEFILES/Ejemplo archivos PSFV XXXX.zip"

dfSociedad = pd.read_csv(rootSociedades)
dfTensiones = pd.read_csv(csvTensiones)
dfTrafos = pd.read_csv(csvTrafos)
dfModulos = pd.read_csv(csvModulos)
dfInverters = pd.read_csv(csvInverters)
dfCeldas = pd.read_csv(csvCeldas)


# Geografía
rootProvincias = "GEOGRAFÍA/Provincias.csv"
rootMunicipios =  "GEOGRAFÍA/Municipios.csv"
dfProvincias = pd.read_csv(rootProvincias)
dfMunicipios = pd.read_csv(rootMunicipios)

# Seleccionar la sociedad
sociedad = st.selectbox("Selecciona la entidad promotora del proyecto:", dfSociedad['NOMBRE'].unique(), 0)

# Diccionario de la sociedad
mainDic['nombreSociedad'] = sociedad
mainDic['CIFSociedad'] = dfSociedad.loc[dfSociedad['NOMBRE']==mainDic['nombreSociedad'],'CIF'].item()
mainDic['direccionSociedad'] = dfSociedad.loc[dfSociedad['NOMBRE']==mainDic['nombreSociedad'],'DIRECCIÓN'].item()
mainDic['logoSociedad'] = dfSociedad.loc[dfSociedad['NOMBRE']==mainDic['nombreSociedad'],'LOGO'].item()
mainDic['logoC'] = mainDic['logoSociedad'] + ".png"
mainDic['logoH'] = mainDic['logoC']

st.divider()


# Seleccionar datos geográficos

col1, col2 = st.columns(2)
with col1:
    provincia = st.selectbox("Selecciona la provincia:", sorted(dfMunicipios['PROVINCIA'].unique()), 0)

with col2:
    dfNumMunic = pd.DataFrame(columns=['Num'])
    dfNumMunic['Num'] = ['Número de municipios'] + list(range(1, 10))
    numberMunic = st.selectbox("Número de municipios", dfNumMunic['Num'])
    if numberMunic != 'Número de municipios':
        listadoMunicipiosProj = []
        for i in range(int(numberMunic)):
            municipioProj = st.selectbox(f"Selecciona el municipio {i+1}",dfMunicipios.loc[dfMunicipios['PROVINCIA']==provincia,'MUNICIPIO'], 0)
            listadoMunicipiosProj.append(municipioProj)
            
        if int(numberMunic) != 1: 
            mainDic['municipioProj'] = ', '.join(listadoMunicipiosProj[:-2]) + ' y ' + listadoMunicipiosProj[-1]
            mainDic['municipioProjC'] = mainDic['municipioProj']
            mainDic['ccaaProj'] = dfMunicipios.loc[dfMunicipios['MUNICIPIO']==listadoMunicipiosProj[-1],'COMUNIDAD AUTÓNOMA'].item()
            mainDic['ccaaProjC'] = mainDic['ccaaProj']
        else:
            mainDic['municipioProj'] = listadoMunicipiosProj[0]
            mainDic['municipioProjC'] = mainDic['municipioProj']
            mainDic['ccaaProj'] = dfMunicipios.loc[dfMunicipios['MUNICIPIO']==municipioProj,'COMUNIDAD AUTÓNOMA'].item()
            mainDic['ccaaProjC'] = mainDic['ccaaProj']



col4, col5 = st.columns(2)

with col4:
    latitud = st.text_input("Latitud del proyecto:")

with col5:
    longitud = st.text_input("Longitud del proyecto:")

areaProj = st.text_input("Área del proyecto [ha]:")

# Diccionario de la geografía
if provincia:
    mainDic['provinciaProj'] = provincia
    mainDic['provinciaProjC'] = provincia


if areaProj:
    mainDic['areaInstalacion'] = normalize(areaProj)

if latitud and longitud:
    mainDic['latitudProj'] = normalize2(latitud)
    mainDic['longitudProj'] = normalize2(longitud)
    mainDic['UTMX'] = str(round(cd.latLonToXY(float(mainDic['latitudProj']), float(mainDic['longitudProj']))[0],2))
    mainDic['UTMY'] = str(round(cd.latLonToXY(float(mainDic['latitudProj']), float(mainDic['longitudProj']))[1],2))
    mainDic['husoUTM'] = str(cd.latLonToXY(float(mainDic['latitudProj']), float(mainDic['longitudProj']))[2])
    mainDic['UTMkey'] = cd.latLonToXY(float(mainDic['latitudProj']), float(mainDic['longitudProj']))[3]

    # Parameters of the cable
    elevacion, presion = cd.get_elevationAndPressure(latitud, longitud)
    elevacion = round(elevacion, 2)
    presion = round(presion, 2)

    mainDic['altitud'] = elevacion #Este a priori no se usa
    mainDic['presion'] = presion
st.divider()

if proyectoTipo == "Fotovoltáico":

    col6, col7 = st.columns(2)

    with col6:
        produccionPlanta = st.text_input("Producción de la planta (MWh/año):")
        if produccionPlanta:
            mainDic['producAnual'] = normalize(produccionPlanta)

    with col7:
        toggle5MW = tog.st_toggle_switch("Menor o igual a 5 MW")
        optionsPV5 = ''
        # Check the toggle button state
        if toggle5MW:
            optionsPV5 = st.text_input("Introduce la potencia instalada EN INVERSORES (MW):")
            mainDic['potProj5'] = "Sí"
            mainDic['tiempoDisparo'] = "0.5"  
            mainDic['mesesProjDur'] = "10"  
            mainDic['mesesProjLetraDur'] = ntl.numero_a_letras(round(float(mainDic['mesesProjDur']))).lower()
            mainDic['figuraCronograma'] = "DATASHEETS/Cronogramas/10meses.png"
            if optionsPV5:
                #Valor para el excelReader
                mainDic['potProjMWac'] = normalize(optionsPV5) #Replace si existe coma, si no, sigue haciendo el round
        else:
            mainDic['potProj5'] = "No"
            mainDic['tiempoDisparo'] = "0.2"            
            mainDic['mesesProjDur'] = "16"  
            mainDic['mesesProjLetraDur'] = ntl.numero_a_letras(round(float(mainDic['mesesProjDur']))).lower()
            mainDic['figuraCronograma'] = "DATASHEETS/Cronogramas/16meses.png"


    # Parámetros de la SET, del cable, del trafo, y de las estructuras

    # Create a radio button with three choices
    st.divider()

    lineaTipo = st.radio("Selecciona el tipo de cable",( "Aéreo", "Subterraneo"))
    mainDic['lineaTipo'] = lineaTipo

    if lineaTipo == "Aéreo":
        
        col8, col9, col10= st.columns(3)
        st.divider()

        nombreSET = st.text_input("Nombre de la SET:")
        tensionSET = st.selectbox("Tensión de la SET", dfTensiones['Un'].unique(), 0)
        
        if nombreSET and (tensionSET != "Un"):
            mainDic['nombreSET'] = nombreSET
            mainDic['tensionSET'] = tensionSET

            # Aprovechamos aquí para coger los datos de las celdas de MT
            mainDic['tensionCeldaMax'] = dfCeldas.loc[dfCeldas['TENSION NOMINAL']== float(mainDic['tensionSET']),'TENSION MAXIMA'].item()
            mainDic['tensionCeldaEnsayo '] = dfCeldas.loc[dfCeldas['TENSION NOMINAL']== float(mainDic['tensionSET']),'TENSION ENSAYO FRECUENCIA'].item()
            mainDic['tensionCeldaCorta'] = dfCeldas.loc[dfCeldas['TENSION NOMINAL']== float(mainDic['tensionSET']),'CORRIENTE ADMISIBLE CORTA DURACION'].item()
            mainDic['tensionCeldaAsignada'] = dfCeldas.loc[dfCeldas['TENSION NOMINAL']== float(mainDic['tensionSET']),'CORRIENTE ASIGNADA EN SERVICIO EMBARRADO'].item()


# mainDic['direccionSociedad'] = dfSociedad.loc[dfSociedad['NOMBRE']==mainDic['nombreSociedad'],'DIRECCIÓN'].item()


        with col8:

            longLinea = st.text_input("Longitud de la línea CT planta --> SET (km):")
            tension = st.selectbox("Tensión de la línea", dfTensiones['Un'].unique(), 0)
            if tension:
                tensionMax = dfTensiones.loc[dfTensiones['Un']==tension,'Up'].item()

        with col9:
    
            voltageDrop = st.slider("Maximum % voltage drop", 0.0, 10.0, 0.0, step  =0.1)
            powerLoss = st.slider("Maximum % power loss", 0.0, 10.0, 0.0, step  =0.1)

        if longLinea and tension and tensionMax and voltageDrop and powerLoss:
            mainDic['tensionAereaLinea'] = tension
            mainDic['tensionMaxAereaLinea'] = tensionMax
            mainDic['longAereaLinea'] = normalize(longLinea)
            mainDic['caidaTension'] = voltageDrop
            mainDic['perdidaPotencia'] = powerLoss
        
        with col10:
            tracker = st.radio("Tipo de estructura:", ("Fija", "Seguidor"))

            if tracker == "Fija":
                tilt = st.text_input("Tilt de la estructura (º)")
                mainDic['seguidorTipo'] = tracker
                mainDic['tiltEstructura'] = tilt
            elif tracker == "Seguidor":
                tiltTracker = st.text_input("Ángulo rotación estructura ±(º)")
                mainDic['trackerTilt '] = tiltTracker
                mainDic['seguidorTipo'] = tracker

                # Esto lo sacamos de si es tipo 1V o demás en readers
                # mainDic['longFilaTracker'] = st.text_input("Longitud de fila por tracker (m)")
                # mainDic['longFilaTracker'] = mainDic['longFilaTracker'] + ' m'
            
            ejes = st.radio("Disposición ejes:", ("Monoposte", "Bi-poste"))

            if ejes:
                mainDic['disposicionEjes'] = ejes

        trafoManuf = st.selectbox("Fabricante del centro de transformación", dfTrafos['MANUFACTURER'].unique(), 0)

        if trafoManuf:
            mainDic['trafoManuf'] = trafoManuf

        st.divider()
        

        dfAux = pd.DataFrame(columns=['Num'])
        dfAux['Num'] = ['Número de municipios'] + list(range(1, 10))
        number = st.selectbox("Número de municipios atravesados por la línea", dfAux['Num'])

        if number != 'Número de municipios':
            listadoMunicipios = []
            for i in range(int(number)):
                municipio = st.text_input(f"Municipio {i+1}")
                listadoMunicipios.append(municipio)
            if int(number) != 1: 
                mainDic['municipiosAtravesados'] = ', '.join(listadoMunicipios[:-2]) + ' y ' + listadoMunicipios[-1]
            else:
                mainDic['municipiosAtravesados'] = listadoMunicipios[0]




# Time to read the documents 

st.divider()

coly, colx = st.columns(2)
with colx:
    uploadedFiles = st.file_uploader("Arrastra los excels y la imagen de la planta (formato png)", accept_multiple_files = True)

with coly:
    st.caption("Ejemplo de archivos")
    with open("SAMPLEFILES/Ejemplo archivos PSFV XXXX.zip", "rb") as fp:
        btn = st.download_button(
            label="Descarga archivos de ejemplo",
            data=fp,
            file_name="Ejemplo archivos PSFV XXXX.zip",
            mime="application/zip"
        )



for uploaded_file in uploadedFiles:
    if uploaded_file.name.endswith("xlsx"):
        if "ProjectSheet" in pd.ExcelFile(uploaded_file).sheet_names:
            pvFile = uploaded_file
            aux_dic = rd.excelReaderPVD(uploaded_file, mainDic, rootEstructuras, dfModulos, user)
            mainDic.update(aux_dic)
            mainDic.update(lineasAereas(mainDic))

        elif (("Vallado") or ("C. V. (Formato Word)")) in pd.ExcelFile(uploaded_file).sheet_names:
            df_parcelasVallado = DataFrame()
            df_parcelasLinea = DataFrame()
            df_centroTrafo = DataFrame()
            df_accesos = DataFrame()

            df_parcelasVallado, df_parcelasLinea, df_centroTrafo, df_accesos = rd.excelReaderCoordenadas(uploaded_file, mainDic)

        elif (("Planta")) in pd.ExcelFile(uploaded_file).sheet_names:
            df_parcelasPlanta = DataFrame()
            df_parcelasTramo = DataFrame()

            #Importing the tables from Viesgo
            df_parcelasPlanta, df_parcelasTramo = rd.excelReaderParcelas(uploaded_file, mainDic)
            
    elif uploaded_file.name.endswith("png"):
        picFile = uploaded_file
        mainDic['tomaAerea'] = picFile.name

    elif uploaded_file.name.endswith("pdf"):

        word = rd.reader(uploaded_file,0)

        if "PVsyst" in word:

            PVsyst_file = uploaded_file

        else:

            planos_file = uploaded_file

        # Datasheets se cogen según los datos

        pass



# Parámetros de los inversores
# st.divider()
try:
    if pvFile:

        # Cambiamos ratioTrafoSET del PVD
        mainDic['ratioTrafoSET'] =  str(mainDic['ratioTrafoSET'].partition('/')[0]) + '/' + mainDic['tensionSET']

        col11, col12 = st.columns(2)
        col13, col14 = st.columns(2)
        col15, col16 = st.columns(2)

        dfAuxTrafos = pd.DataFrame(columns=['Num'])
        dfAuxTrafos['Num'] = ['Número de transformadores'] + list(range(1, 20))


        with col11: #Numero de transformadores
            numberTrafos = st.selectbox("Número de transformadores", dfAuxTrafos['Num'])
            mainDic['numTrafos'] = numberTrafos

            if (str(numberTrafos) != '1') and (numberTrafos != "Número de transformadores"):
                toggle_button_trafos = tog.st_toggle_switch("Misma potencia en todos los trafos", "hey", label_after=False)     

        with col12: #Potencia de los transformadores
            if str(numberTrafos) == '1': #Sólo un trafo
                trafoUnico = st.text_input("Potencia Transformador (MVA)")
                mainDic['potTrafos'] = trafoUnico
                mainDic['trafosPlanta'] = str(mainDic['numTrafos'])  + " transformador de potencia de " + mainDic['potTrafos'] + " MVA." 

                if float(mainDic['numInverter']) > 1:
                    mainDic['trafoLongTab'] = str(mainDic['numTrafos'])  + " transformador de " + mainDic['potTrafos'] + " MVA que verterá la energía producida por " + str(mainDic['numInverter']) + " inversores."
                else:
                    mainDic['trafoLongTab'] = str(mainDic['numTrafos'])  + " transformador de " + mainDic['potTrafos'] + " MVA que verterá la energía producida por " + str(mainDic['numInverter']) + " inversor."

                mainDic['trafoTab'] = str(mainDic['numTrafos'])  + " x " + mainDic['potTrafos']

            elif (toggle_button_trafos) and (numberTrafos != 'Número de transformadores en la planta'): #Todos los trafos con la misma potencia
                trafoUnico = st.text_input("Potencia Transformador (MVA)")
                #Varios trafos con la misma potencia
                mainDic['potTrafos'] = trafoUnico
                mainDic['trafosPlanta'] = str(mainDic['numTrafos'])  + " transformadores de potencia de " + mainDic['potTrafos'] + " MVA cada uno." 

                if float(mainDic['numInverter']) > 1:
                    mainDic['trafoLongTab'] = str(mainDic['numTrafos'])  + " transformador de " + mainDic['potTrafos'] + " MVA que verterá la energía producida por " + str(mainDic['numInverter']) + " inversores."
                else:
                    mainDic['trafoLongTab'] = str(mainDic['numTrafos'])  + " transformador de " + mainDic['potTrafos'] + " MVA que verterá la energía producida por " + str(mainDic['numInverter']) + " inversor."

                mainDic['trafoTab'] = str(mainDic['numTrafos'])  + " x " + mainDic['potTrafos']
                
            elif numberTrafos != 'Número de transformadores': #Varios trafos con distintas potencias
                listadoPotenciaTrafos = []
                for i in range(int(numberTrafos)):
                    trafo = st.text_input(f"Potencia Transformador (MVA) {i+1}")
                    listadoPotenciaTrafos.append(trafo)

                trafosClean = list(set(listadoPotenciaTrafos)) #Todas las potencias que hay

                if trafosClean:
                    mainDic['potTrafos'] = ', '.join(trafosClean[:-2]) + ' y ' + trafosClean[-1]


                mainDic['trafosPlanta'] = str(mainDic['numTrafos'])  + " transformadores de potencia de " + mainDic['potTrafos'] + " MVA cada uno." #Que si solo pone uno, que no ponga cada uno
                if float(mainDic['numInverter']) > 1:
                    mainDic['trafoLongTab'] = str(mainDic['numTrafos'])  + " transformadores de " + mainDic['potTrafos'] + " MVA que verterá la energía producida por " + str(mainDic['numInverter']) + " inversores."
                else: 
                    mainDic['trafoLongTab'] = str(mainDic['numTrafos'])  + " transformadores de " + mainDic['potTrafos'] + " MVA que verterá la energía producida por " + str(mainDic['numInverter']) + " inversor."

        with col13: #Fabricante de los inversores
            inverterManuf = st.selectbox("Fabricante de los inversores a potencia máxima", sorted(dfInverters['MANUFACTURER'].unique()), 0)
            mainDic['inverterManufNonL'] = inverterManuf
       
        with col14: #Modelo de los inversores
            inverterModel = st.selectbox("Modelo inversores a potencia máxima", sorted(dfInverters.loc[dfInverters['MANUFACTURER']==inverterManuf, 'MODEL']), 0)
            mainDic['inverterModelNonL'] = inverterModel

        #Valores de los inversores no limitados
        mainDic['rangoUNonL'] = dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'RANGO DE TENSIÓN EN MPP'].item()
        mainDic['UMaxNonL'] = dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'TENSIÓN MÁXIMA'].item()
        mainDic['IMaxNonL'] = dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'CORRIENTE MÁXIMA POR ENTRADA'].item()
        mainDic['entradasDCNonL'] = dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'ENTRADAS EN DC'].item()
        mainDic['PnNonL'] = dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'POTENCIA NOMINAL'].item()
        mainDic['UnNonL'] = dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'TENSIÓN NOMINAL'].item()
        mainDic['fnNonL'] = '50'
        mainDic['etaMaxNonL'] = dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'RENDIMIENTO MÁXIMO'].item()
        mainDic['etaEuNonL'] = dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'RENDIMIENTO EUROPEO'].item()

        mainDic['inverterNonLimit'] = mainDic['numInverter']
        mainDic['inverterLimitText'] = ", con una potencia máxima de " + str(mainDic['PnNonL']) + " kW."
        mainDic['potInverter'] = mainDic['PnNonL']
        mainDic['diverseInvertersInfo'] = "del fabricante " + mainDic['inverterManufNonL'] + " , modelo " + mainDic['inverterModelNonL']
        mainDic['stringCent'] = str(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'TIPO'].item())
        mainDic['shortcircPow'] = str(round(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'Pk'].item()*int(mainDic['numInverter']),2))
        if mainDic['stringCent'] == "String":
              mainDic['stringCent'] = " en String"      

        numInversores = int(mainDic['numInverter'])
        potInverter = float(mainDic['PnNonL'])
        toggleInverter = tog.st_toggle_switch("Inversores con potencia limitada")


        if toggleInverter: #Existen inversores con la potencia limitada

            with col15:
                st.divider()
                numInvCapados = st.selectbox(":red[Inversores con la potencia limitada]",list(range(1,numInversores+1)), 0)
                inverterManufCapados = st.selectbox(":red[Fabricante de los inversores a potencia limitada]", sorted(dfInverters['MANUFACTURER'].unique()), 0)                
            with col16:
                st.divider()
                potInvCapados = st.text_input(":red[Potencia máxima de los inversores limitados (kVA)]")
                if (potInvCapados) and (float(potInvCapados) > potInverter):
                    st.write("La potencia tiene que ser menor que " + str(potInverter) + "kVA")       
                inverterModelCapados = st.selectbox(":red[Modelo inversores a potencia limitada]", sorted(dfInverters.loc[dfInverters['MANUFACTURER']==inverterManuf, 'MODEL']), 0)

            # Valores en el caso de que haya inversores limitados
            if inverterManufCapados and inverterModelCapados:
                mainDic['inverterLManuf'] = inverterManufCapados
                mainDic['inverterLModel'] = inverterModelCapados

            mainDic['stringCent'] = str(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'TIPO'].item())
            mainDic['stringHCent'] =  mainDic['stringCent'] #.style => blueHeader


            mainDic['rangoUL'] = str(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterLModel'],'RANGO DE TENSIÓN EN MPP'].item())
            mainDic['UMaxL'] = str(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterLModel'],'TENSIÓN MÁXIMA'].item())
            mainDic['IMaxL'] = str(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterLModel'],'CORRIENTE MÁXIMA POR ENTRADA'].item())
            mainDic['entradasDCL'] = str(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterLModel'],'ENTRADAS EN DC'].item())
            if potInvCapados:
                mainDic['PnL'] = potInvCapados
            mainDic['UnL'] = str(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterLModel'],'TENSIÓN NOMINAL'].item())
            mainDic['fnL'] = '50'
            mainDic['etaMaxL'] = str(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterLModel'],'RENDIMIENTO MÁXIMO'].item())
            mainDic['etaEuL'] = str(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterLModel'],'RENDIMIENTO EUROPEO'].item())
            if numInvCapados and numInversores:
                mainDic['invertersLimit'] = numInvCapados
                mainDic['invertersNonLimit'] = str(numInversores - numInvCapados)
            mainDic['shortcircPow'] = str(round(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterModelNonL'],'Pk'].item()*int(mainDic['invertersNonLimit']) + dfInverters.loc[dfInverters['MODEL']==mainDic['inverterLModel'],'Pk'].item()*int(mainDic['invertersLimit']),2)) 

            mainDic['stringCent'] = str(dfInverters.loc[dfInverters['MODEL']==mainDic['inverterLModel'],'TIPO'].item())
            if potInvCapados and potInverter:
                mainDic['potInverter'] = str(mainDic['PnNonL'])+'/'+str(mainDic['PnL'])

                mainDic['inverterLimitText'] = ", siendo " + str(mainDic['invertersNonLimit']) + " de ellos a potencia máxima de " + str(mainDic['PnNonL']) + " kW y " + str(mainDic['invertersLimit']) + " de ellos con limitación a "  + str(mainDic['PnL'])+ "kW ."

            if inverterManuf == inverterManufCapados:
                mainDic['diverseInvertersInfo'] = "del fabricante " + mainDic['inverterManufNonL'] + ", modelo " + mainDic['inverterModelNonL']
            else:
                mainDic['diverseInvertersInfo'] = "del fabricante " + mainDic['inverterManufNonL'] + ", modelo " + mainDic['inverterModelNonL'] + "y del fabricante " + mainDic['inverterLManuf'] + ", modelo " + mainDic['inverterLModel']

            if mainDic['stringCent'] == "String":
                mainDic['stringCent'] = " de String"

except NameError:
    pass

st.divider()

# Duplicamos documentos
try:
    if lineaTipo and tracker:
        if lineaTipo == "Aéreo":
            if tracker == "Fija":
                try:
                    if toggleInverter:
                        fileModelo = "Modelo anteproyecto AFL.docx"
                except Exception as e:
                    pass
                else:
                    fileModelo = "Modelo anteproyecto AFNL.docx"
            else:
                try:
                    if toggleInverter:
                        fileModelo = "Modelo anteproyecto ATL.docx"
                except Exception as e:
                    pass                
                else:
                    fileModelo = "Modelo anteproyecto ATNL.docx"
        elif lineaTipo == "Subterraneo":
            if tracker == "Fija":
                try:
                    if toggleInverter:
                        fileModelo = "Modelo anteproyecto SFL.docx"
                except Exception as e:
                    pass                
                else:
                    fileModelo = "Modelo anteproyecto SFNL.docx"
            else:
                try:
                    if toggleInverter:
                        fileModelo = "Modelo anteproyecto STL.docx"
                except Exception as e:
                    pass                
                else:
                    fileModelo = "Modelo anteproyecto STNL.docx"

    fileAnexoCalculos = "Modelo Portada Calculos Energeticos.docx"
    fileAnexoEquipos = "Modelo Portada Catalogo Equipos.docx"
    fileAnexoPlanos = "Modelo Portada Planos.docx"

    numEdicion = st.text_input("Edición del documento:")
    mainDic['versionDoc'] = numEdicion
    mainDic['versionDocC'] = numEdicion
    flagGenerar = st.button("Generar documento")

    # Preparamos los documentos de los anejos, arriba ya deberíamos tener el de planos y el del PVsyst (planos_file y PVsyst_file). Falta hacer merge de los datasheet según los parámetros escogidos

    if ejes:
        if ejes == "Bi-poste":
            pass
            # "DATASHEETS/Estructuras/monolineplus-1p-datasheet.pdf"



    if flagGenerar and fileModelo:
        with st.spinner("Generando documento"):
            # Estilos especiales
            #Cover Bold
            mainDic['potPicoC'] = mainDic['potPico']
            mainDic['potInstaladaC'] = mainDic['potInstalada']
            mainDic['nombreProyectoC'] = mainDic['nombreProyecto']

            mainDic['FlagReference'] = ''
            mainDic['FlagFigRef'] = ''

            #Cover Light

            mainDic['dateCoverC'] = month.capitalize()

            # Fechas
            mainDic['dateMY'] = month
            mainDic['dateCoverC'] = month.capitalize()
            mainDic['date'] = dt.datetime.now().strftime("%d/%m/%y")

            fileModelo =  fileModelo
            # fileModelo = rootWord + "\\" + fileModelo  #Esto si lo teníamos dentro de una carpeta no funcionaba. ¿Puede ser que usar \\ en vez de / era lo que fallase? @note 

            doc_modelo, doc_anexoCalculos, doc_anexoEquipos, doc_anexoPlanos = duplicateDoc(fileModelo), duplicateDoc(fileAnexoCalculos),duplicateDoc(fileAnexoEquipos), duplicateDoc(fileAnexoPlanos)

            flagPicWriter = insert_image_in_cell(doc = doc_modelo, dic = mainDic, picFile = picFile)
            flagPicWriterCalculos, flagPicWriterEquipos, flagPicWriterPlanos = insert_image_in_cell(doc = doc_anexoCalculos, dic = mainDic, picFile = picFile), insert_image_in_cell(doc = doc_anexoEquipos, dic = mainDic, picFile = picFile), insert_image_in_cell(doc = doc_anexoPlanos, dic = mainDic, picFile = picFile)
            flagPicList = [flagPicWriter, flagPicWriterCalculos, flagPicWriterEquipos, flagPicWriterPlanos]
 

            flagDocWriter = wt.docWriter(doc_modelo,mainDic)
            flagDocWriterCalculos, flagDocWriterEquipos, flagDocWriterPlanos = wt.docWriter(doc_anexoCalculos,mainDic), wt.docWriter(doc_anexoEquipos,mainDic), wt.docWriter(doc_anexoPlanos,mainDic)

            flagDocList = [flagDocWriter, flagDocWriterCalculos, flagDocWriterEquipos, flagDocWriterPlanos]


            flagDocTabler = wt.docTabler(doc_modelo,df_parcelasPlanta, df_parcelasTramo,df_parcelasVallado, df_centroTrafo, df_parcelasLinea, df_accesos)
            flagDocTabler = [flagDocTabler]

            checkList = flagPicList + flagDocList + flagDocTabler


            if all(element == 1 for element in checkList):
                import io
                doc_modelo_bio = io.BytesIO()
                doc_modelo.save(doc_modelo_bio)
                doc_modelo_bio.seek(0)
                nameDocModelo = 'AyC ' + mainDic['nombreProyecto'] + " " + "Ed." + mainDic['versionDoc'] + ".docx"

                doc_anexoCalculos_bio = io.BytesIO()
                doc_anexoCalculos.save(doc_anexoCalculos_bio)
                doc_anexoCalculos_bio.seek(0)
                nameDocAnexoCalculos = 'Anexo Calculos AyC ' + mainDic['nombreProyecto'] + " " + "Ed." + mainDic['versionDoc'] + " " + ".docx"

                doc_anexoEquipos_bio = io.BytesIO()
                doc_anexoEquipos.save(doc_anexoEquipos_bio)
                doc_anexoEquipos_bio.seek(0)
                nameDocAnexoEquipos = 'Anexo Equipos AyC ' + mainDic['nombreProyecto'] + " " + "Ed." + mainDic['versionDoc'] + " " + ".docx"

                doc_anexoPlanos_bio = io.BytesIO()
                doc_anexoPlanos.save(doc_anexoPlanos_bio)
                doc_anexoPlanos_bio.seek(0)
                nameDocAnexoPlanos = 'Anexo Planos AyC ' + mainDic['nombreProyecto'] + " " + "Ed." + mainDic['versionDoc'] + " " + ".docx"

                nameZip = 'AyC ' + mainDic['nombreProyecto'] + " " + "Ed." + str(mainDic['versionDoc']) + " " + ".zip"
                zip_data = io.BytesIO()

                # Create a ZipFile Object
                with ZipFile(zip_data, 'w') as zipf:
                   # Adding files that need to be zipped
                    zipf.writestr(nameDocModelo, doc_modelo_bio.getvalue())
                    zipf.writestr(nameDocAnexoCalculos,doc_anexoCalculos_bio.getvalue())
                    zipf.writestr(nameDocAnexoEquipos, doc_anexoEquipos_bio.getvalue())
                    zipf.writestr(nameDocAnexoPlanos, doc_anexoPlanos_bio.getvalue())
                    flagZip = 1
                    flagZip = [flagZip]
                    checkList = checkList + flagZip
                

        if all(element == 1 for element in checkList):
            st.info('¡Documento listo para descarga!')
            btn = st.download_button(
                label="Descarga archivos",
                data=zip_data.getvalue(),
                file_name=nameZip,
                mime="application/zip"
            )

except Exception as e:
    st.write(e)   
