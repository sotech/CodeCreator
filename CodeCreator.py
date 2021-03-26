#Class code creator
#Input excel file
import xlrd
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from decimal import Decimal, getcontext, ROUND_HALF_UP
import os
import json

round_context = getcontext()
round_context.rounding = ROUND_HALF_UP

#Hardcoded States and Companies
stateNames = {"Alabama" : "AL",
              "Alaska" : "AK",
              "Arizona" : "AZ",
              "Arkansas" : "AR",
              "California" : "CA",
              "Colorado" : "CO",
              "Connecticut" : "CT",
              "Delaware" : "DE",
              "District of Columbia" : "DC",
              "Florida" : "FL",
              "Georgia" : "GA",
              "Hawaii" : "HI",
              "Idaho" : "ID",
              "Illinois" : "IL",
              "Indiana" : "IN",
              "Iowa" : "IA",
              "Kansas" : "KS",
              "Kentucky" : "KY",
              "Louisiana" : "LA",
              "Maine" : "ME",
              "Maryland" : "MD",
              "Massachusetts" : "MA",
              "Michigan" : "MI",
              "Minnesota" : "MN",
              "Mississippi" : "MS",
              "Missouri" : "MO",
              "Montana" : "MT",
              "Nebraska" : "NE",
              "Nevada" : "NV",
              "New Hampshire" : "NH",
              "New Jersey" : "NJ",
              "New Mexico" : "NM",
              "New York" : "NY",
              "North Carolina" : "NC",
              "North Dakota" : "ND",
              "Ohio" : "OH",
              "Oklahoma" : "OK",
              "Oregon" : "OR",
              "Pennsylvania" : "PA",
              "Rhode Island" : "RI",
              "South Carolina": "SC",
              "South Dakota" : "SD",
              "Tennessee" : "TN",
              "Texas" : "TX",
              "Utah" : "UT",
              "Vermont" : "VT",
              "Virginia" : "VA",
              "Washington" : "WA",
              "West Virginia" : "WV",
              "Wisconsin" : "WI",
              "Wyoming" : "WY"}

stateCities = {"Alabama" : "Acmar",
              "Alaska" : "Ambler",
              "Arizona" : "Aguila",
              "Arkansas" : "Alicia",
              "California" : "Afton",
              "Colorado" : "Alma",
              "Connecticut" : "Avon",
              "Delaware" : "Bethel",
              "District of Columbia" : "Anacostia",
              "Florida" : "Alford",
              "Georgia" : "Adel",
              "Hawaii" : "Anahola",
              "Idaho" : "Almo",
              "Illinois" : "Adams",
              "Indiana" : "Akron",
              "Iowa" : "Albia",
              "Kansas" : "Agenda",
              "Kentucky" : "Ajax",
              "Louisiana" : "Akers",
              "Maine" : "Amity",
              "Maryland" : "Allen",
              "Massachusetts" : "Agawam",
              "Michigan" : "Alba",
              "Minnesota" : "Akeley",
              "Mississippi" : "Ansley",
              "Missouri" : "Alexandria",
              "Montana" : "Alzada",
              "Nebraska" : "Alvo",
              "Nevada" : "Austin",
              "New Hampshire" : "Andover",
              "New Jersey" : "Alloway",
              "New Mexico" : "Abbott",
              "New York" : "Acra",
              "North Carolina" : "Ahoskie",
              "North Dakota" : "Adrian",
              "Ohio" : "Adelphi",
              "Oklahoma" : "Aline",
              "Oregon" : "Alameda",
              "Pennsylvania" : "Adah",
              "Rhode Island" : "Bradford",
              "South Carolina": "Alcolu",
              "South Dakota" : "Alsen",
              "Tennessee" : "Algood",
              "Texas" : "Abram",
              "Utah" : "Alta",
              "Vermont" : "Athens",
              "Virginia" : "Aldie",
              "Washington" : "Acme",
              "West Virginia" : "Advent",
              "Wisconsin" : "Alvin",
              "Wyoming" : "Alva"}

stateZips = {"Alabama" : "35004",
              "Alaska" : "99786",
              "Arizona" : "85320",
              "Arkansas" : "72410",
              "California" : "95920",
              "Colorado" : "80420",
              "Connecticut" : "06001",
              "Delaware" : "19931",
              "District of Columbia" : "20373",
              "Florida" : "32420",
              "Georgia" : "31620",
              "Hawaii" : "96703",
              "Idaho" : "83312",
              "Illinois" : "62347",
              "Indiana" : "46910",
              "Iowa" : "52531",
              "Kansas" : "66930",
              "Kentucky" : "41722",
              "Louisiana" : "70421",
              "Maine" : "04471",
              "Maryland" : "21810",
              "Massachusetts" : "01001",
              "Michigan" : "49611",
              "Minnesota" : "56433",
              "Mississippi" : "39558",
              "Missouri" : "63430",
              "Montana" : "59311",
              "Nebraska" : "68304",
              "Nevada" : "89310",
              "New Hampshire" : "03216",
              "New Jersey" : "08001",
              "New Mexico" : "87747",
              "New York" : "12405",
              "North Carolina" : "27910",
              "North Dakota" : "58472",
              "Ohio" : "43101",
              "Oklahoma" : "73716",
              "Oregon" : "97211",
              "Pennsylvania" : "15410",
              "Rhode Island" : "02808",
              "South Carolina": "29001",
              "South Dakota" : "57004",
              "Tennessee" : "38501",
              "Texas" : "78572",
              "Utah" : "84092",
              "Vermont" : "05143",
              "Virginia" : "20105",
              "Washington" : "98220",
              "West Virginia" : "25231",
              "Wisconsin" : "54542",
              "Wyoming" : "82711"}

companies = ("A6","CS","L7","K2","M8","M9","MX","I9")

#Clase Informacion
class Informacion:
    def __init__(self,_companyCode,_effDate,_city,_state,_stateName,_stateZip,_insuredName,_entity,_experienceMod,_scheduleMod,_increasedLimits):
        Informacion.companyCode = _companyCode
        Informacion.effDate = _effDate
        Informacion.city = _city
        Informacion.state = _state
        Informacion.stateName = _stateName
        Informacion.stateZip = _stateZip
        Informacion.insuredName = _insuredName
        Informacion.entity = _entity
        Informacion.experienceMod = _experienceMod
        Informacion.scheduleMod = _scheduleMod
        Informacion.increasedLimits = _increasedLimits

class Codigo:
    def __init__(self,_codigo,_rate,_premium):
        Codigo.codigo = _codigo
        Codigo.rate = _rate
        Codigo.premium = _premium
        
def FormateoCodigo(codigo):
    codigoFormateado = codigo
    while len(codigoFormateado) < 4:
            codigoFormateado = "0" + codigoFormateado
    return codigoFormateado

#Funciones
def CrearArchivo(information, _vcodes, _vrates, _vpremium, _vtags, _vexposure, _filename):
    #Run X amount of times according to the limit
    tam = len(_vcodes)
    print("Creando archivo...")
    data = {}
    data["CompanyCode"] = information.companyCode
    data["EffDate"] = information.effDate
    data["Entity"] = information.entity
    data["ExperienceModifier"] = information.experienceMod
    data["IncreasedLimits"] = information.increasedLimits
    data["InsuredCity"] = information.city
    data["InsuredName"] = information.insuredName
    data["InsuredState"] = information.state
    data["InsuredStateName"] = information.stateName
    data["InsuredZip"] = information.stateZip
    data["ScheduleModifier"] = information.scheduleMod
    #Adicion de codigos, rates y exposure
    #Codigos normales
    premiumTotal = 0
    for i in range(tam):
        #Exposure
        if i == 0:
            campoexposure = _vtags[i]
            campocode = "ClassCode"
            camporate = "Rate"
        else:
            campoexposure = _vtags[i] + str(i)
            campocode = "ClassCode" + str(i)
            camporate = "Rate" + str(i)
        data[campoexposure] = str(_vexposure[i])
        #Code
        codigo = FormateoCodigo(str(int(_vcodes[i])))        
        data[campocode] = codigo
        #Premium
        premiumTotal = premiumTotal + int(_vpremium[i])
        #Rate
        data[camporate] = _vrates[i]
    data["TotalManualPremium"] = int(premiumTotal)
    datos = []
    datos.append(data)
    with open(_filename + ".json", 'w') as outfile:
        json.dump(datos, outfile, ensure_ascii = False, indent = 4)
    print("Archivo creado")

def CalcularPremium(rate):
    return int(round(rate * 100))

def CalcularRate(lossCost, lcm):
    producto = lcm * lossCost
    return c_round(producto,2)

def c_round(x, digits, precision=5):
    tmp = round(Decimal(x), precision)
    return float(tmp.__round__(digits))

def Import():
    filepath = filedialog.askopenfilename(initialdir = os.getcwd(), title = "Select Class Codes Excel file to import", filetypes=(("XLSX","*.xlsx"),("XLS","*.xls")))
    lb_filename.config(text=filepath)
    
def Extraccion(filename,lcm):
    #Inicializacion
    codes = xlrd.open_workbook(filename)
    hoja = codes.sheet_by_index(0)

    vcodes = []
    vrates = []
    vpremium = []
    vexposureTag = []
    vexposurePremium = []

    vcodesdis = []
    vratesdis = []
    vpremiumdis = []
    vexposureTagdis = []
    vexposurePremiumdis = []

    vcodesnon = []
    vratesnon = []
    vpremiumnon = []
    vexposureTagnon = []
    vexposurePremiumnon = []

    vcodescapita = []
    vratescapita = []
    vpremiumcapita = []
    vexposureTagcapita = []
    vexposurePremiumcapita = []

    #Extraccion
    for i in range(hoja.nrows):
        if hoja.cell_value(i,2) != "–" and hoja.cell_value(i,2) != "-" and ("a" not in str(hoja.cell_value(i,2))):
            #No es discontinuo, no es "a"
            if ("N" not in hoja.cell_value(i,1)):
                #No es non-ratable
                if ("P" not in hoja.cell_value(i,1)):
                    #No es per capita
                    vcodes.append(hoja.cell_value(i,0))
                    valor = hoja.cell_value(i,2)
                    print(valor)
                    rate = CalcularRate(float(hoja.cell_value(i,2)),lcm)
                    vrates.append(rate)
                    vpremium.append(CalcularPremium(rate))
                    vexposureTag.append("AnnualExposure")
                    vexposurePremium.append(10000)
                else:
                    #Es per capita
                    vcodescapita.append(hoja.cell_value(i,0))
                    rate = CalcularRate(float(hoja.cell_value(i,2)),lcm)
                    vratescapita.append(rate)
                    vpremiumcapita.append(CalcularPremium(rate))
                    vexposureTagcapita.append("PerCapitaAnnualExposure")
                    vexposurePremiumcapita.append(100)
            else:
                #Es non-ratable
                vcodesnon.append(hoja.cell_value(i,0))
                rate = CalcularRate(float(hoja.cell_value(i,2)),lcm)
                vratesnon.append(rate)
                vexposureTagnon.append("AnnualExposure")
                vexposurePremiumnon.append(10000)
                if hoja.cell_value(i,3) != "–":
                    #Non ratable calculable                    
                    vpremiumnon.append(CalcularPremium(rate))
                else:
                    #Non ratable no calculable
                    vpremiumnon.append("0")
        else:
            if ("a" not in str(hoja.cell_value(i,2))):
                #Es discontinuo
                vcodesdis.append(hoja.cell_value(i,0))
                vratesdis.append("0")
                vpremiumdis.append("0")
                vexposureTagdis.append("AnnualExposure")
                vexposurePremiumdis.append(0)
                
    #Hay per capita? Agregar un normal code para su ejecucion
    if len(vcodescapita) > 0 :
        vcodescapita.insert(0,vcodes[0])
        vratescapita.insert(0,vrates[0])
        vpremiumcapita.insert(0,vpremium[0])
        vexposureTagcapita.insert(0,vexposureTag[0])
        vexposurePremiumcapita.insert(0,vexposurePremium[0])

    #Hay discontinuos? Agregar un normal code para su ejecucion
    if len(vcodesdis) > 0 :
        vcodesdis.insert(0,vcodes[0])
        vratesdis.insert(0,vrates[0])
        vpremiumdis.insert(0,vpremium[0])
        vexposureTagdis.insert(0,vexposureTag[0])
        vexposurePremiumdis.insert(0,vexposurePremium[0])
    
    datos = {}
    datos['normc'] = vcodes
    datos['normr'] = vrates
    datos['normp'] = vpremium
    datos['normt'] = vexposureTag
    datos['norme'] = vexposurePremium
    datos['percac'] = vcodescapita
    datos['percar'] = vratescapita
    datos['percap'] = vpremiumcapita
    datos['percat'] = vexposureTagcapita
    datos['percae'] = vexposurePremiumcapita
    datos['nonc'] = vcodesnon
    datos['nonr'] = vratesnon
    datos['nonp'] = vpremiumnon
    datos['nont'] = vexposureTagnon
    datos['none'] = vexposurePremiumnon
    datos['disc'] = vcodesdis
    datos['disr'] = vratesdis
    datos['disp'] = vpremiumdis
    datos['dist'] = vexposureTagdis
    datos['dise'] = vexposurePremiumdis
    datos['tam'] = hoja.nrows
    
    
    return datos

def AnadirNotificacion(mensaje):
    lb_not.config(text=lb_not.cget("text") + mensaje)
    
def Notificacion(mensaje):
    lb_not.config(text=mensaje)
    
def Generate():
    #Extraer inputs
    statename = txt_stfl.get()
    state = stateNames[statename]
    file = lb_filename.cget("text")
    if txt_lcm.get() == "":
        lcm = float(1)
    else:
        lcm = float(txt_lcm.get())
    companycode = txt_cc.get()
    effdate = txt_ed.get()
    city = stateCities[statename]
    statezip = stateZips[statename]
    insuredname = "Test Insured"
    entity = "Policy"
    experiencemod = "1"
    schedule = "1"
    increasedlimits = "1000/1000/1000"
    if txt_lim.get() != "":
        limite = int(txt_lim.get())
    else:
        limite = 0
    informacion = Informacion(companycode,effdate,city,state,statename,statezip,insuredname,entity,experiencemod,schedule,increasedlimits)
    #Extraer Codigos
    datos = Extraccion(file,lcm)
    #Crear Archivos
    #Subdivision de codigos
    #Si la cantidad supera el limite pedido, dividirlos
    #Normales
    tam = datos['tam']
    if len(datos['normc']) > limite and limite != 0:
        i = 0
        v = 0
        nombreBase = "normalCodes"
        vectorC = []
        vectorR = []
        vectorP = []
        vectorT = []
        vectorE = []
        for code in datos['normc']:
            vectorC.append(code)
            vectorR.append(datos['normr'][i + v * limite])
            vectorP.append(datos['normp'][i + v * limite])
            vectorT.append(datos['normt'][i + v * limite])
            vectorE.append(datos['norme'][i + v * limite])
            if i >= (limite-1):
                i = 0
                lb_not.config(text="Generating Normal Codes " + str(v+1) + "...\n")
                nombreArchivo = nombreBase + str(v+1)
                CrearArchivo(informacion,vectorC,vectorR,vectorP,vectorT,vectorE,nombreArchivo)
                AnadirNotificacion("Done!\n")
                v = v + 1
                vectorC = []
                vectorR = []
                vectorP = []
                vectorT = []
                vectorE = []
            else:
                i = i + 1
        #Revisar si sobran
        if len(vectorC) > 0:
            #Hay codigos sobrantes, hacer un archivo con esos
            Notificacion("Generating Remaining Normal Codes...\n")
            nombreArchivo = nombreBase + str(v+1)
            CrearArchivo(informacion,vectorC,vectorR,vectorP,vectorT,vectorE,nombreArchivo)
            AnadirNotificacion("Done!\n")
    else:
        AnadirNotificacion("Generating Normal Codes...\n")
        CrearArchivo(informacion,datos['normc'],datos['normr'],datos['normp'],datos['normt'],datos['norme'],"normalCodes")
        AnadirNotificacion("Done!\n")
    if len(datos['percac']) > 0:
        AnadirNotificacion("Generating PerCapita Codes...\n")
        CrearArchivo(informacion,datos['percac'],datos['percar'],datos['percap'],datos['percat'],datos['percae'],"percapitaCodes")
        AnadirNotificacion("Done!\n")
    else:
        AnadirNotificacion("No per capita codes found!\n")

    if len(datos['nonc']) > 0:
        AnadirNotificacion("Generating Non Ratable Codes...\n")
        CrearArchivo(informacion,datos['nonc'],datos['nonr'],datos['nonp'],datos['nont'],datos['none'],"nonratableCodes")
        AnadirNotificacion("Done!\n")
    else:
        AnadirNotificacion("No Non Ratable Codes found!\n")

    if len(datos['disc']) > 0:
        AnadirNotificacion("Generating Discontinued Codes...\n")
        CrearArchivo(informacion,datos['disc'],datos['disr'],datos['disp'],datos['dist'],datos['dise'],"discontinuedCodes")
        AnadirNotificacion("Done!")
    else:
        AnadirNotificacion("No Discontinued Codes found!\n")


#Ventana
ventana = Tk()
ventana.title("Code Creator v2.1")
ventana.geometry("500x400")
#CompanyCode
lb_cc = Label(ventana, text="Company Code(*): ")
lb_cc.grid(column=0, row=1)
txt_cc = Combobox(ventana)
txt_cc['values'] = companies
txt_cc.current(1)
txt_cc.grid(column=1,row=1)
#LCM
lb_lcm = Label(ventana, text="LCM. (Default 1): ")
lb_lcm.grid(column=0,row=2)
txt_lcm = Entry(ventana, width=10)
txt_lcm.grid(column=1,row=2)
#Effective date
lb_ed = Label(ventana, text="Effective date MM/DD/YYYY(*): ")
lb_ed.grid(column=0, row=3)
txt_ed = Entry(ventana, width=10)
txt_ed.grid(column=1,row=3)
#State full
lb_stfl = Label(ventana, text="State(*): ")
lb_stfl.grid(column=0, row=4)
txt_stfl = Combobox(ventana)
statesFull = list(stateNames.keys())
txt_stfl['values'] = statesFull
txt_stfl.grid(column=1,row=4)
#Limite
lb_limite = Label(ventana, text = "Limit codes per file. (Default ALL): ")
lb_limite.grid(column=0, row=5)
txt_lim = Entry(ventana,width=10)
txt_lim.grid(column=1,row=5)
#Required
lb_required = Label(ventana, text = "Required Field (*)")
lb_required.grid(column=0, row=6)
#File
lb_file = Label(ventana, text="Select the file to import: ")
lb_file.grid(column=0, row=9)
btn_file = Button(ventana, text="Import file", command=Import)
btn_file.grid(column=1,row=9)
lb_filename = Label(ventana, text="")
lb_filename.grid(column=0,row=10)
#Generate
btn_gen = Button(ventana, text="Generate files", command=Generate)
btn_gen.grid(column=0,row=11)
#Notifications
lb_not = Label(ventana, text="")
lb_not.grid(column=0,row=12)
ventana.mainloop()


