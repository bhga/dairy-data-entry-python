import PySimpleGUI as sg
import pandas as pd
import copy
from datetime import date

# Adding colour to the window
sg.theme("DarkTeal9")

Output_path = "Data.xlsx" #this is the file path in which we will keep appending our data

df = pd.read_excel(Output_path) #we read the excel data and created the Dataframe of it

layout = [
    [sg.Text("Dairy Name",size=(20,1)),sg.Combo(["My Dairy","Our Dairy","Your Hotel"],key="Dairy_Name")], #input feilds
    
    [sg.Text("XYZ Milk Delivered(500ml)",size=(20,1)),sg.InputText(key="XYZ_500ml")],

    [sg.Text("ABC Milk Delivered(1000ml)",size=(20,1)),sg.InputText(key="ABC_1000ml")],

    [sg.Text("QRS Milk Delivered(6000ml)",size=(20,1)),sg.InputText(key="QRS_6000ml")],

    [sg.Text("Amount Received",size=(20,1)),sg.InputText(key="Amount_Received")],
    
    [sg.Submit(),sg.Button('Clear'),sg.Exit()] #Event button 
    ]

myWindow = sg.Window("My Milk Collection",layout) #Window variable

def clear_input(): #function to clear screen data
    for key in values:
        myWindow[key]('')
    return None

while True: #while true is used to keep the window alive 
    event,values = myWindow.read() #all the activity that happens in window is stored in event and values
    if event == sg.WIN_CLOSED or event == "Exit": #if user closes the window or clicks exit we break out of the while loop 
        break
    if event == 'Clear': #if event is clear then call the function to clear screen
        clear_input()
    if event == "Submit": #if event is submit then append the data into excel and clear the screen
        
        new_values = copy.deepcopy(values) #we create deep copy of the user values dict
        #note: why we created deep copy cause if we want to add new key to the values then that is assumed as the new feild which is not in window
        #and then will throw us an error while clearing the fields
        
        #Calculating expected amount based on filter condition based on dairy name we are applying specific rate to milk brand
        #why different rates cause there might a discount for golden client and also default values for others
        if new_values["Dairy_Name"] == "My Dairy":
            new_values["Expected_Amount"] = int(new_values["XYZ_500ml"]) * 20 + int(new_values["ABC_1000ml"]) * 19 + int(new_values["QRS_6000ml"]) * 20
        elif new_values["Dairy_Name"] == "Our Dairy":
            new_values["Expected_Amount"] = int(new_values["XYZ_500ml"]) * 20 + int(new_values["ABC_1000ml"]) * 19 + int(new_values["QRS_6000ml"]) * 19
        elif new_values["Dairy_Name"] == "Your Hotel":
            new_values["Expected_Amount"] = int(new_values["XYZ_500ml"]) * 20 + int(new_values["ABC_1000ml"]) * 19 + int(new_values["QRS_6000ml"]) * 18
        else:
            new_values["Expected_Amount"] = int(new_values["XYZ_500ml"]) * 20 + int(new_values["ABC_1000ml"]) * 20 + int(new_values["QRS_6000ml"]) * 20

        #calculating balance amount 
        new_values["Today_Balance"] = int(new_values["Expected_Amount"]) - int(new_values["Amount_Received"])

        #adding insert_date
        new_values["Insert_date"] = date.today()
        
        df_values=pd.DataFrame(new_values,index=[0])#converting dict into dataFrame
        df = pd.concat([df,df_values],axis=0) #appending new data into existing df i.e our excel file

        df["Insert_date"] = pd.to_datetime(df["Insert_date"]).dt.date #keeping all our date in same format
        
        df.to_excel(Output_path,index=False)
        sg.popup("Loaded Successfully!!!")#once data is loaded into excel popup the msg
        clear_input()
myWindow.close() # we close the window
