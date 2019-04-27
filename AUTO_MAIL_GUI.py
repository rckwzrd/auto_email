import PySimpleGUI as sg

form = sg.FlexForm('POTATO MAILER')
form_rows = [
			[sg.Text('Well Info:', size=(20, 1), font=(25), text_color='blue')],
			[sg.Text('Well Name:', size = (12,1)), sg.InputText('Name used to format subject and save .msg')],
			[sg.Text('UWI:', size = (12,1)), sg.InputText('Identifier to used to query contacts in SQL DB')],
			[sg.Text('Ops Geo:', size = (12,1)), sg.InputText('Name used to format email signature')],
			[sg.Text('To:', size = (12,1)), sg.InputText('Pulled from SQL DB')],
			[sg.Text('CC:', size = (12,1)), sg.InputText('Manual Entry')],
			# Find screenshots and save folder
			[sg.Text('Attachments and Folders:', size=(20, 1), font=(25), text_color='blue')],
          	[sg.Text('Geo Model:', size=(12, 1)), sg.InputText('Find.jpg screen shot'), sg.FolderBrowse()],
          	[sg.Text('Well Compare:', size=(12, 1)), sg.InputText('Find .jpg screen shot'), sg.FolderBrowse()],
			[sg.Text('Well Folder:', size=(12, 1)), sg.InputText('Directory to save update email'), sg.FolderBrowse()],
			# Get geosteering info
			[sg.Text('Geosteering Info:', size=(20, 1), font=(25), text_color='blue')],
			[sg.Text('Last MD:', size=(12, 1)), sg.InputText('xxxx',size = (5,1))], 
			[sg.Text('New MD:', size=(12, 1)), sg.InputText('xxxx', size = (5,1))],
			[sg.Text('Rig Status:', size=(12, 1)), sg.InputText('DRL LAT', size = (10,1))],
			[sg.Text('Bit Projection:', size = (12,1)), sg.InputText('Interpreted bit position')],
			[sg.Text('Structure:', size = (12,1)), sg.InputText('Expected formation structure')],
			[sg.Text('Carbonates:', size = (12,1)), sg.InputText('Carbonate hazards')],
			[sg.Text('Target:', size = (12,1)), sg.InputText('Optimal drilling interval')],
			[sg.Text('Tolerances:', size = (12,1)), sg.InputText('Lateral tolerances')],
			# Make email
			[sg.Submit('Create Update'), sg.Cancel()]
			]
values = form.LayoutAndRead(form_rows)



