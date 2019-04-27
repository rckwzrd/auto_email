####################################
### AUTO MAILER V3, Matt Rhoads  ###
####################################

import win32com.client as win32
import csv 
import datetime as dt
import pyodbc

#################
### FUNCTIONS ###
#################

def query_wv(uwi):
    '''Connects to SQL DB with well number and returns a list of rig emails.'''
    cur_date = dt.date.today()
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=;DATABASE=;Trusted_Connection=True;')
    cursor = conn.cursor()
    cursor.execute(
    "select * FROM (select email, contactname, title,h.wellname,ROW_NUMBER() OVER(PARTITION BY c.title ORDER BY c.sysmoddate desc) as rownum from wvt_wvjobcontact c inner join wvt_wvwellheader h on c.idwell = h.idwell  where  wellidb = '"+uwi+"' and email is not null) T where T.rownum = 1")
    x = ''
    rows = cursor.fetchall()
    for row in rows:
        x += row[0] + "; "
    print(x)
    return x

def get_time_day():
    '''Get time of day and format to string with elif statements. This is called in rig_format().'''
    hour = dt.datetime.today().hour
    t = 'null'
    if 0 <= hour <= 11:
        t = 'Geo update for the morning.'
    elif 12 <= hour <= 17:
        t = 'Geo update for the afternoon.'
    elif 18 <= hour <= 24:
        t = 'Geo update for the evening.'
    return t

def rig_stat(s):
    '''Format rig status to a string based on rig_status variable. This is called in rig_format().'''
    print(s)
    s_text= ""
    if s == "DRL_LAT":
        s_text = "Rig is currently drilling ahead in lateral"
    elif s == "TIH":
        s_text = "Rig is currently TIH"
    elif s == "TOOH":
        s_text = "Rig is currently TOOH"
    elif s == "DRL_CURVE":
        s_text = "Rig is currently buidling curve"
    return s_text
 

def rig_format(new_md, old_md, rig_status, bit_projection, structure, target, carbaonte, tolerances, survey):
    '''Use input variables to format the subject and body of an HTML email and return in a list.'''
    time = get_time_day()
    rig_text = rig_stat(rig_status)
    dist_drl = str(int(new_md) - int(old_md))
    intro = time+" "+dist_drl+"'MD drilled since last report. See comments below and attached geosteering model. Please call if you have questions."
    subject = new_md+'MD|'+well_name+'|Geosteering Update'
    body = '<html> <body style="font-family:Arial; font-size:12pt">' 
    body += intro
    body += "<br><br><b>As of last survey at "+new_md+"'MD:</b>"
    body += "<ul><li><b><u>Rig Status:</b></u> "+rig_text +"</li>"
    body += "<li><b><u>Bit Projection:</b></u> "+bit_projection+"</li>"
    body += "<li><b><u>Expected Structure:</b></u> "+structure+"</li>"
    body += "<li><b><u>Target:</b></u> "+target+"</li>"
    body += "<li><b><u>Carbonate:</b></u> "+carbonate+"</li>"
    body += '<li><u><font color ="red"><b>Lateral Tolerances:</b></u> '+tolerances+'</font></li></ul>'
    body += "<br>"+survey+"<br>"
    body += "<br>Matt Rhoads<br>"
    body += "Operations Geologist<br>"
    body += '</body> </html>'
    the_list = [subject,body]
    return the_list


def rig_mail():
    '''Take control of outlook, create email, assign recipients, attach screen shots, save, and send.'''
    rig_mail = outlook.createItem(0)
    rig_mail.To = to_contact
    rig_mail.CC = cc_contact
    rig_mail.Subject = rig_text[0]
    rig_mail.HtmlBody = rig_text[1]
    rig_mail.Attachments.Add(gm_path)
    rig_mail.Attachments.Add(wc_path) 
    rig_mail.Display(False)
    rig_mail.SaveAs(save_file_name)

##############
### INPUTS ###
##############

# Rig contacts pulled from sql query: needs to be refined
uwi = ""
wv_contact = query_wv(uwi)

# Direct rig contacts: Drilling engineer, CM, DD, MWD, and Well Clerk
to_contact = ""

# Additional contacts: Area superintendant, geo superviser, drilling supervisor, and well drive auto upload
cc_contact = ""


# well, folder, and geo name
well_name = ""
folder_name = ""
geo_name = "RhoadsM"

# current and old MD of bit
new_md = "11000"
old_md = "10000"

# Path for geosteering screen shots and save email 
gm_path = 'C:.../Screenshots/.jpg'
wc_path = 'C:.../Screenshots/.jpg'
save_file_name = 'C:/Users/rhoad515/Desktop/Auto_Mail/'+folder_name+'/'+new_md+'MD_'+well_name+'_GEOSTEERING_UPDATE.msg'

# Enter rig status and geosteering comments for each update here.
# DRL_LAT, TIH, TOOH, DRL_CURVE
rig_status = "TIH"
bit_projection = "Bit is at 'TVD, ' below TOT, ' above BOT, and ' below plan."
structure = "Formation is expected to dip. Plan dip is parallel to formation dip."
target = "Recommend targeting >80API GR shale ' below plan line."
carbonate = "20-60API GR carbonate stringer is presentt ' below plan line "
tolerances = " 0' up / 25' down relative to plan #4."
survey = ""

######################
### GENERATE EMAIL ###
######################

# Create outlook object
outlook = win32.Dispatch('outlook.application')

# Feed geosteering comments to rig_format to make subject and body
rig_text = rig_format(new_md, old_md, rig_status, bit_projection, structure, target, carbonate, tolerances, survey)

# Make email, assign contacts, fill in subject/body, attach screen shots, save, and send
rig_mail()


