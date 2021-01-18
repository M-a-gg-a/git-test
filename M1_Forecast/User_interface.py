# Module 1: user interface

from tkinter import Tk, Label, Button, StringVar, OptionMenu, Entry, IntVar, Checkbutton, END, LabelFrame


root = Tk()
root.geometry("1920x1080")
root.title("Surplus Interconnection Service")


#functions
# 2a function for transmission line utilization
def utilization_data():
    button_utilization.config(state = 'disabled')
    utilization_input.config(state = 'disabled')
    
    utilization = 0    
    utilization = float(utilization_input.get())
    if utilization > 1:
        utilization = [utilization/100]                 # variable stores the percentage to simulate in a list       
    else:
        utilization = [float(utilization_input.get())]  # variable stores the percentage to simulate as list
    
    print("utilization factor:", utilization)
    # temporary display
    #myLabel = Label(root, text = utilization, font = 'consolas 11').grid(row = 13, column = 4, sticky = 'E')

# 3a function for the period to simulate the renewable resources        
def year_data():
    button_years.config(state = 'disabled')
    years_input.config(state = 'disabled')
        
    year = 0
    year = [int(years_input.get())]                     # variable stores years to simulate as list
    print("years to simulate: ",year)
    # temporary display
    #myLabel = Label(root, text = year, font = 'consolas 11').grid(row = 17, column = 4, sticky = 'E')

# 4a function factor data for any adjusting factor needed
def factor_data():
    button_factor.config(state = 'disabled')
    factor_input.config(state = 'disabled')
    
    factor = 0
    factor = [float(factor_input.get())]                # variable factor stores the value as list
    print("adjusting factor to simulate: ", factor)
    # temporary display
    #myLabel = Label(root, text = factor, font = 'consolas 11').grid(row = 21, column = 4, sticky = 'E')

#5a fuction to let the user to upload a gas forward curve, the code to do the process needs to be added
def gasforwardcurve():
    button_gasforwardcurve.config(state = 'disabled')
    return
#6a fuction to let the user to upload a gas forward curve, the code to do the process needs to be added
def marketprice():
    button_gasforwardcurve.config(state = 'disabled')
    return
 

# 0 main title
main_title = Label(root, text= "Welcome to Surplus Interconnection Service Simulator", font = 'Arial 14 bold')
main_title.grid(row = 5, column = 1, columnspan = 5, pady = 40, sticky = 'E')


# 1. Select transmission line to simulate based on a list of existing locations with MW nameplate
#drop down list for locations
loc_simulate = Label(root, text= "Select location to simulate: ", font = 'Arial 12')
loc_simulate.grid(row = 9, column = 2, columnspan = 2, padx = 100, pady = 15, sticky = 'W')

options = ["Carousel : wind",
           "San Isabel : solar",
           "Twin Buttes : wind",
           "Colorado Highlands Wind : wind",
           "Crossing Trails Wind Farm : wind",
           "Kit Carson Wind Farm : wind",
           "New TS resource : wind"                     #list of existing renewable resources
]

#Definition of variable location, dropdown menu and set initial location to Carousel
loc = StringVar()                                # variable loc stores selected location as list 
loc.set(options[0])
drop = OptionMenu(root, loc, *options).grid(row = 9, column = 4, pady = 15, sticky = 'W')


# 2 user enters the percentage of transmission line to simulate 
# it calls the function utilization, stores the percentage to simulate in the variable utilization as list
uti_simulate = Label(root, text= "Enter % of transmission line to simulate: ", font = 'Arial 12')
uti_simulate.grid(row = 13,  column = 3, padx = 100, pady = 15, sticky = 'W')

utilization_input = Entry(root, width= 30, borderwidth=5, font = 'Arial 9 bold')
utilization_input.grid(row = 13, column = 4, pady = 15, sticky = 'W')

button_utilization = Button(root, text = "Apply", font = 'Arial 10 bold', command = utilization_data)
button_utilization.grid(row = 13, column = 5, padx = 2, pady =15, sticky = 'W')

    
# 3 user enters how many years to simulate, it calls the function year data
# the number of years is store in the variable year as list
years_simulate = Label(root, text = "Enter how many years to simulate: ", font = 'Arial 12')
years_simulate.grid(row = 17, column = 3, padx = 100, pady = 15, sticky = 'W')

years_input = Entry(root, width= 30, borderwidth = 5, font = 'Arial 9 bold')
years_input.grid(row = 17, column = 4, pady = 15, sticky = 'W')

button_years = Button(root, text = "Apply", font = 'Arial 10 bold', command = year_data)
button_years.grid(row = 17, column = 5, pady =15, sticky = 'W')


# 4 user enters any adjusting factor if needed, it calls the function factor_data
# the factor is store in the variable factor as list
factor_simulate = Label(root, text = "Enter any adjusting factor to simulate: ", font = 'Arial 12', justify = 'left')
factor_simulate.grid(row = 21, column = 3, padx = 100, pady = 15, sticky = 'W')

factor_input = Entry(root, width= 30, borderwidth=5, font = 'Arial 9 bold')
factor_input.grid(row = 21, column = 4, pady = 15, sticky = 'W')

button_factor = Button(root, text = "Apply", font = 'Arial 10 bold', command = factor_data)
button_factor.grid(row = 21, column = 5, pady =15, sticky = 'W')


#5 the user has the option to upload a forward curve csv file, the code to do the process needs
#the process calls the functin gasforwardcurve but does nothing yet
upload_gasforwardcurve = Label(root, text = "Upload forward gas csv file:  ", font = 'Arial 12')
upload_gasforwardcurve.grid(row = 17, column = 6, pady = 15, sticky = 'W')

button_gasforwardcurve = Button(root, text = "Submit", font = 'Arial 10 bold', command = gasforwardcurve)
button_gasforwardcurve.grid(row = 17, column = 7, padx = 2, pady =15, sticky = 'W')


#6 the user has the option to upload a forward curve csv file, the code to do the process needs to be added
#the process calls the functin gasforwardcurve but does nothing yet
upload_marketprice = Label(root, text = "Upload a market price csv file: ", font = 'Arial 12')
upload_marketprice.grid(row = 21, column = 6,  pady = 15, sticky = 'W')

button_marketprice = Button(root, text = "Submit", font = 'Arial 10 bold', command = marketprice)
button_marketprice.grid(row = 21, column = 7, padx = 2, pady =15, sticky = 'W')


#7 the user is given 20 different options of stand-alone or combination of resources to simulate
#the selected values are store in the renewable variable as list
renewable_simulate = Label(root, text = "Select renewable resources to simulate: ", font = 'Arial 12 bold')
renewable_simulate.grid(row = 27, column = 3, columnspan=5, padx = 100, pady = 35, sticky = 'W')


#7a function selected that stores the values of the renewable resources selected to be simulated and cleans screen values
def selected():
    #myLabel = Label(root, text = loc.get()).grid(row = 9, column = 5, padx = 50, pady = 15, sticky = 'E')
    myButton.config(state = 'disabled')
    wind_solar_checkb.config(state = 'disabled')
    wind_CC_checkb.config(state = 'disabled')
    wind_CT_checkb.config(state = 'disabled')
    wind_RIC_checkb.config(state = 'disabled')
    wind_Aro_checkb.config(state = 'disabled')
    wind_solar_batt_checkb.config(state = 'disabled')
    #wind_solar_CC_checkb.config(state = 'disabled')
    #wind_solar_CT_checkb.config(state = 'disabled')
    #wind_solar_RIC_checkb.config(state = 'disabled')
    #wind_solar_Aro_checkb.config(state = 'disabled')
    solar_wind_checkb.config(state = 'disabled')
    solar_CC_checkb.config(state = 'disabled')
    solar_CT_checkb.config(state = 'disabled')
    solar_RIC_checkb.config(state = 'disabled')
    solar_Aro_checkb.config(state = 'disabled')
    solar_wind_batt_checkb.config(state = 'disabled')
    solar_wind_CC_checkb.config(state = 'disabled')
    solar_wind_CT_checkb.config(state = 'disabled')
    solar_wind_RIC_checkb.config(state = 'disabled')
    solar_wind_Aro_checkb.config(state = 'disabled')
   
    
    location = [loc.get()]
    print("location to simulate:", location)
    renewables = [wind_solar.get(), wind_CC.get(), wind_CT.get(), wind_RIC.get(),
                  wind_Aro.get(), wind_solar_batt.get(),
                  solar_wind.get(), solar_CC.get(), solar_CT.get(), solar_RIC.get(),
                  solar_Aro.get(), solar_wind_batt.get(), solar_wind_CC.get(), solar_wind_CT.get(), solar_wind_RIC.get(), solar_wind_Aro.get(),
                  solar_wind_RIC.get(), solar_wind_Aro.get(),  
                  ] #stores selected resources
    
    print("renewable resources to simulate:", '\n', renewables)    
         
#Definition of variables for the renewable checkbutton resources
wind_solar = IntVar()
wind_CC = IntVar()
wind_CT = IntVar()
wind_RIC = IntVar()
wind_Aro = IntVar()
wind_solar_batt = IntVar()
#wind_solar_CC = IntVar()
#wind_solar_CT = IntVar()
#wind_solar_RIC = IntVar()
#wind_solar_Aro = IntVar()

solar_wind = IntVar()
solar_CC = IntVar()
solar_CT = IntVar()
solar_RIC = IntVar()
solar_Aro = IntVar()
solar_wind_batt = IntVar()
solar_wind_CC = IntVar()
solar_wind_CT = IntVar()
solar_wind_RIC = IntVar()
solar_wind_Aro = IntVar()


#renewable resources checkbuttons
wind_solar_checkb = Checkbutton(root, text = "wind - solar", font = 'Arial 11', variable = wind_solar, borderwidth=4)
wind_solar_checkb.grid(row = 28, column = 3, pady = 2, sticky = 'E')
wind_CC_checkb = Checkbutton(root, text = "wind - CC  ", variable = wind_CC, font = 'Arial 11', borderwidth=4)
wind_CC_checkb.grid(row = 32, column = 3, pady = 2, sticky = 'E')
wind_CT_checkb = Checkbutton(root, text = "wind - CT   ", variable = wind_CT, font ='Arial 11', borderwidth=4)
wind_CT_checkb.grid(row = 36, column = 3, pady = 2, sticky = 'E')
wind_RIC_checkb = Checkbutton(root, text = "wind - RIC ", variable = wind_RIC, font = 'Arial 11', borderwidth=4)
wind_RIC_checkb.grid(row = 40, column = 3, pady = 2, sticky = 'E')
wind_Aro_checkb = Checkbutton(root, text = "wind - Aro  ", variable = wind_Aro, font = 'Arial 11', borderwidth=4)
wind_Aro_checkb.grid(row = 28, column = 4, padx = 75, pady = 2, sticky = 'W')
wind_solar_batt_checkb = Checkbutton(root, text = "wind - solar - batteries", variable = wind_solar_batt, font = 'Arial 11', borderwidth=4)
wind_solar_batt_checkb.grid(row = 32, column = 4, padx = 75, pady = 2, sticky = 'W')
#wind_solar_CC_checkb = Checkbutton(root, text = "wind - solar - CC", variable = wind_solar_CC, font = 'Arial 11', borderwidth=4)
#wind_solar_CC_checkb.grid(row = 32, column = 4, pady = 2, sticky = 'W')
#wind_solar_CT_checkb = Checkbutton(root, text = "wind - solar - CT", variable = wind_solar_CT, font = 'Arial 11', borderwidth=4)
#wind_solar_CT_checkb.grid(row = 36, column = 4, pady = 2, sticky = 'W')
#wind_solar_RIC_checkb = Checkbutton(root, text = "wind - solar - RIC", variable = wind_solar_RIC, font = 'Arial 11', borderwidth=4)
#wind_solar_RIC_checkb.grid(row = 40, column = 4, pady = 2, sticky = 'W')
#wind_solar_Aro_checkb = Checkbutton(root, text = "wind - solar - Aro", variable = wind_solar_Aro, font = 'Arial 11', borderwidth=4)
#wind_solar_Aro_checkb.grid(row = 44, column = 4, pady = 2, sticky = 'W')

solar_wind_checkb = Checkbutton(root, text = "solar - wind", font = 'Arial 11', variable = solar_wind, borderwidth=4)
solar_wind_checkb.grid(row = 36, column = 4, padx = 75, pady = 2, sticky = 'W')
solar_CC_checkb = Checkbutton(root, text = "solar - CC  ", variable = solar_CC, font = 'Arial 11', borderwidth=4)
solar_CC_checkb.grid(row = 40, column = 4, padx = 75, pady = 2, sticky = 'W')
solar_CT_checkb = Checkbutton(root, text = "solar - CT   ", variable = solar_CT, font ='Arial 11', borderwidth=4)
solar_CT_checkb.grid(row = 28, column = 5, pady = 2, sticky = 'W')
solar_RIC_checkb = Checkbutton(root, text = "solar - RIC ", variable = solar_RIC, font = 'Arial 11', borderwidth=4)
solar_RIC_checkb.grid(row = 32, column = 5, pady = 2, sticky = 'W')
solar_Aro_checkb = Checkbutton(root, text = "solar - Aro  ", variable = solar_Aro, font = 'Arial 11', borderwidth=4)
solar_Aro_checkb.grid(row = 36, column = 5, pady = 2, sticky = 'W')
solar_wind_batt_checkb = Checkbutton(root, text = "solar - wind - batteries", variable = solar_wind_batt, font = 'Arial 11', borderwidth=4)
solar_wind_batt_checkb.grid(row = 40, column = 5, pady = 2, sticky = 'W')
solar_wind_CC_checkb = Checkbutton(root, text = "solar - wind - CC", variable = solar_wind_CC, font = 'Arial 11', borderwidth=4)
solar_wind_CC_checkb.grid(row = 28, column = 6, padx = 25, pady = 2, sticky = 'W')
solar_wind_CT_checkb = Checkbutton(root, text = "solar - wind - CT", variable = solar_wind_CT, font = 'Arial 11', borderwidth=4)
solar_wind_CT_checkb.grid(row = 32, column = 6, padx = 25, pady = 2, sticky = 'W')
solar_wind_RIC_checkb = Checkbutton(root, text = "solar - wind - RIC", variable = solar_wind_RIC, font = 'Arial 11', borderwidth=4)
solar_wind_RIC_checkb.grid(row = 36, column = 6, padx = 25, pady = 2, sticky = 'W')
solar_wind_Aro_checkb = Checkbutton(root, text = "solar - wind - Aro", variable = solar_wind_Aro, font = 'Arial 11', borderwidth=4)
solar_wind_Aro_checkb.grid(row = 40, column = 6, padx = 25, pady = 2, sticky = 'W')


#8 Button to submit selected renewables resources, it calls the function selected 
myButton = Button(root, text = "Submit", font ='Arial 11 bold', command = selected)
myButton.grid(row = 60, column = 4, columnspan = 2, padx = 130, pady = 20, sticky = 'W')

#9 next trial
def another_trial():
     # Delete submitted utilization, years and factor information in the entry boxes
        
        
    loc.set(options[0])
    button_utilization.config(state = 'normal')
    button_years.config(state = 'normal')
    button_factor.config(state = 'normal')
    button_gasforwardcurve.config(state = 'normal')
    button_gasforwardcurve.config(state = 'normal')
    myButton.config(state = 'normal')
    
    utilization_input.config(state = 'normal')
    utilization_input.delete(0, END)
    
    years_input.config(state = 'normal')
    years_input.delete(0, END)
    
    factor_input.config(state = 'normal')
    factor_input.delete(0, END)
    
    
     
    wind_solar_checkb.config(state = 'normal')
    wind_CC_checkb.config(state = 'normal')
    wind_CT_checkb.config(state = 'normal')
    wind_RIC_checkb.config(state = 'normal')
    wind_Aro_checkb.config(state = 'normal')
    wind_solar_batt_checkb.config(state = 'normal')
    #wind_solar_CC_checkb.config(state = 'normal')
    #wind_solar_CT_checkb.config(state = 'normal')
    #wind_solar_RIC_checkb.config(state = 'normal')
    #wind_solar_Aro_checkb.config(state = 'normal')
    solar_wind_checkb.config(state = 'normal')
    solar_CC_checkb.config(state = 'normal')
    solar_CT_checkb.config(state = 'normal')
    solar_RIC_checkb.config(state = 'normal')
    solar_Aro_checkb.config(state = 'normal')
    solar_wind_batt_checkb.config(state = 'normal')
    solar_wind_CC_checkb.config(state = 'normal')
    solar_wind_CT_checkb.config(state = 'normal')
    solar_wind_RIC_checkb.config(state = 'normal')
    solar_wind_Aro_checkb.config(state = 'normal')
    
    
    
    wind_solar_checkb.deselect()
    wind_CC_checkb.deselect()
    wind_CT_checkb.deselect()
    wind_RIC_checkb.deselect()
    wind_Aro_checkb.deselect()
    wind_solar_batt_checkb.deselect()
    #wind_solar_CC_checkb.deselect()
    #wind_solar_CT_checkb.deselect()
    #wind_solar_RIC_checkb.deselect()
    #wind_solar_Aro_checkb.deselect()
    solar_wind_checkb.deselect()
    solar_CC_checkb.deselect()
    solar_CT_checkb.deselect()
    solar_RIC_checkb.deselect()
    solar_Aro_checkb.deselect()
    solar_wind_batt_checkb.deselect()
    solar_wind_CC_checkb.deselect()
    solar_wind_CT_checkb.deselect()
    solar_wind_RIC_checkb.deselect()
    solar_wind_Aro_checkb.deselect()    
 

try_again_Button = Button(root, text = "Try again", font ='Arial 11 bold', command = another_trial)
try_again_Button.grid(row = 60, column = 6, columnspan = 2, padx = 130, pady = 20, sticky = 'W')



root.mainloop()
