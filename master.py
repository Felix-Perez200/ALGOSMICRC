import pandas as pd
from datetime import datetime,date
import xlsxwriter
import datetime as dt
import numpy as np
import PySimpleGUI as sg
import traceback
import os

    


def my_function(filePath,filePath_INT,filePath2):

    filePath = filePath.strip("‪u202a")
    df=df2=df3=df4=df5=pd.read_excel (r'{}'.format(filePath))

    filePath_INT=filePath_INT.strip("‪u202a")
    df_int= pd.read_excel (r'{}'.format(filePath_INT))

    #--------------------------------------------------------------Dataframe Acumulados----------------------------------------------------------------
    df = df.rename(columns={'createdAtDate': 'fecha' })
    df['fecha'] = pd.to_datetime(df['fecha']).dt.date
    # condition0 = df['patient_address_city'] == 'Arecibo'
    # df = df[condition0]
    df = df.groupby(['fecha']).caseType.value_counts().unstack(fill_value=0)
    sum_column = df["Confirmed"] + df["Probable"] + df["Suspected"]
    df["Total"] = sum_column
   
    df['Acumulativo(Confirmado)'] = df['Confirmed'].cumsum()
    df = df.rename(columns={'Confirmed': 'Confirmado','Suspected': 'Sospechoso' })

    #----------------------------------Dataframe hospitalizados-----------------------------------------------------------------------------------------------
    df2['patient_birthDate'] = pd.to_datetime(df2['patient_birthDate']).dt.date
    condition = df2['isCurrentlyHospitalized'] == True
    new_dataframe = df2[condition]
    df2 =new_dataframe.filter(["patient_birthDate","patient_sex", "patient_address_city","caseType","caseStatus"])
    ref_date = dt.datetime.now()
    df2['age'] = df2['patient_birthDate'].apply(lambda x: len(pd.date_range(start = x, end = ref_date, freq = 'Y'))) 
    df2 = df2[["patient_birthDate", "age","patient_sex", "patient_address_city","caseType","caseStatus"]]
    df2 = df2.rename(columns={'patient_birthDate': 'DOB','age': 'EDAD','patient_sex':'Genero','patient_address_city':'Ciudad','caseType':'Tipo','caseStatus':'Estado' })
    df2['Reportado']=np.nan

    #---------------------------------Dataframe Viajeros-----------------------------------------------------------------------------------------
    df3['patient_birthDate'] = pd.to_datetime(df3['patient_birthDate']).dt.date
    df3.fillna(0, inplace = True)
    condition = df3['traveledDuringLast14Days'] == True
    df3 = df3[condition]
    df3 =df3.filter(["caseId","patient_birthDate","patient_sex", "patient_address_city","caseType","caseStatus"])
    ref_date = dt.datetime.now()
    df3['age'] = df3['patient_birthDate'].apply(lambda x: len(pd.date_range(start = x, end = ref_date, freq = 'Y'))) 
    df3 = df3[["caseId","patient_birthDate", "age","patient_sex", "patient_address_city","caseType","caseStatus"]]
    df3 = df3.rename(columns={'patient_birthDate': 'DOB','age': 'EDAD','patient_sex':'Genero','patient_address_city':'Ciudad','caseType':'Tipo','caseStatus':'Estado' })
    df3['Reportado']=np.nan

    #---------------------------------------------------Dataframe GIS---------------------------------------------------------------------------
    df4.fillna(0, inplace = True)
    ref_date = dt.datetime.now()
    df4['age'] = df4['patient_birthDate'].apply(lambda x: len(pd.date_range(start = x, end = ref_date, freq = 'Y')))
    df4 =df4.filter(["patient_birthDate","age","patient_sex", "patient_address_city","caseType","caseStatus"])
    df4 = df4.rename(columns={'patient_birthDate': 'Fecha de nacimiento','age': 'EDAD','patient_sex':'Genero','patient_address_city':'Ciudad','caseType':'Tipo','caseStatus':'Estado' })
    df4['Reportado']=np.nan

    #--------------------------------------------------Dataframe Incidencia-----------------------------------------------------------------------------------
    column_names2 =[  'Fecha',
                    'Casos positivos (PCR)',
                    '14 Dias',
                    'Habitantes en el pueblo',
                    'x 100,000'
    ]
    df6 = pd.DataFrame(columns = column_names2)

    day = dt.datetime.today().strftime("%m/%d/%Y")
    df_last_14 = df['Confirmado'].tail(14).sum()
    amount = 14
    habitantes = 81966
    incidencia = ((df_last_14/amount)/habitantes)*100000

    new_row = {'Fecha':day, 'Casos positivos (PCR)': df_last_14, '14 Dias':amount,'Habitantes en el pueblo': habitantes, 'x 100,000':incidencia}
    df6 = df6.append(new_row, ignore_index=True)
    #--------------------------------DATAFRAME OFICIAL---------------------------------------------------------------------------------------------
    column_names =[  'Día del reporte',
                    'Tasa de incidencia',
                    'Tasa de positividad (últimos 14 días)',
                    'Hospitalizados',
                    'Viajeros',
                    'Muertes',
                    'Confirmados acumulados',
                    'Confirmados adicionales',
                    'Probables acumulados',
                    'Probables adicionales',
                    'Sospechoso Acumulados',
                    'Sospechoso adicionales',
                    'Cerrados acumulados',
                    'Entrevistas realizadas',
                    'Contactos identificados',
                    'Sintomático',
                    'Asintomático',
                    'Hombres',
                    'Mujeres',
                    'Casos activos',
                    'Hospitalizados Activos',
                    'Viajeros Activos']
    df_Oficial = pd.DataFrame(columns = column_names)

    

    dia= dt.datetime.today().strftime("%m/%d/%Y")
    tasa = incidencia
    hosp = df5.shape[0] - df5.lastHospitalizedDate.str.contains(r'0001-01-01T00:00:00Z').sum() 
    death = df5.patient_deceased.sum()
    confirmed = df5.caseType.str.contains(r'Confirmed').sum() 
    sus = df5.caseType.str.contains(r'Suspected').sum()
    prob = df5.caseType.str.contains(r'Probable').sum()
    close = df5.caseStatus.str.contains(r'Closed').sum()
    contacts  = df5['numberOfExposedPositiveContacts'].sum() 
    boy = df5.patient_sex.str.contains(r'Male').sum()
    girl = df5.patient_sex.str.contains(r'Female').sum()
    status = df5.caseStatus.str.contains(r'Active').sum()
    Hosp_status = df5.isCurrentlyHospitalized.sum()
    travel = df5.traveledDuringLast14Days.sum()

    #-Reads from second file(interview DATASET)_
    entrevistas  = df_int.interview_preInterview_correctPerson.str.contains(r'Yes',r'Yes Guardian').sum()
    temp = df_int.drop_duplicates(subset='patientId', keep="first")                 #Remueve dublicados y lo asigna a un dataframe nuevo 
    symptoms= temp.interview_pastSymptoms_hasSymptoms.str.contains(r'Symptomatic').sum()
    asymptoma = temp.interview_pastSymptoms_hasSymptoms.str.contains(r'Asymptomatic').sum()

    new_row = {'Día del reporte':dia,'Tasa de incidencia':tasa , 'Hospitalizados': hosp, 'Muertes': death, 'Confirmados acumulados':confirmed,'Sospechoso Acumulados': sus,
    'Probables acumulados':prob, 'Cerrados acumulados':close, 'Entrevistas realizadas':entrevistas, 'Contactos identificados':contacts, 'Sintomático':symptoms, 'Asintomático':asymptoma, 'Hombres': boy,'Mujeres': girl, 'Mujeres': girl, 'Casos activos': status, 'Hospitalizados Activos':Hosp_status, 'Viajeros Activos':travel}

    df_Oficial = df_Oficial.append(new_row, ignore_index=True)

    #-------------------------------------------------------------------------------------------------------------------------------------------
    writer = pd.ExcelWriter(r'{}.xlsx'.format(filePath2), engine='xlsxwriter')

    # Position the dataframes in the worksheet.
    df.to_excel(writer, sheet_name='Acumulados')  # Default position, cell A1.
    df4.to_excel(writer, sheet_name='Gis',index=False)
    df2.to_excel(writer, sheet_name='Hozpitalizados',index=False)
    df3.to_excel(writer, sheet_name='Viajeros',index=False)
    df_Oficial.to_excel(writer, sheet_name='Oficial',index=False)
    df6.to_excel(writer, sheet_name='Incidencia',index=False)

    writer.save()


def main():
 

    sg.theme("DarkTeal2")

    layout = [
            [sg.T("")], 
            [sg.Text("Choose a casesDataset file: "), sg.Input(), sg.FileBrowse(key="-IN-")],
            [sg.Text("Choose a interviewsDataset file: "),sg.Input(), sg.FileBrowse(key="-IN2-")],
            [sg.Text("Save As: "),sg.Input(), sg.FileSaveAs(key="-OUT-")],
            [sg.Button("Submit")]
            ]

    window = sg.Window('My File Browser', layout, size=(800,250))
        
    try:
        while True:
            event, values = window.read()
             
            if event == sg.WIN_CLOSED or event=="Exit":
                break
            if event == "Submit": 
                for i in range(30):
                    sg.PopupAnimated(sg.DEFAULT_BASE64_LOADING_GIF, message='Please Wait...', background_color="white", text_color='black',time_between_frames=100)                   
                my_function(values['-IN-'],values['-IN2-'],values['-OUT-'])
                sg.PopupAnimated(None)

        window.close()
            
    except Exception as e:
        tb = traceback.format_exc()
        sg.Print(f'An error happened.  Here is the info:', e, tb)
        sg.popup_error(f'AN EXCEPTION OCCURRED!', e, tb)

if __name__ == "__main__":
                    main()
