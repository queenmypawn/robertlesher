# To work with the file system
import sys, win32com.client
import copy
import re
import os

# To work with the .csv
import pandas as pd

# Visio constants
visCharacterColor  = 1
visCharacterFont = 0
visSectionCharacter = 3
visCharacterSize = 7
visCharacterDblUnderline = 8
visSectionFirstComponent = 10

visSectionObject =  1 
visRowPrintProperties =  25 

visPrintPropertiesPageOrientation =  16 
visRowPage =  10 
visPageWidth =  0 
visPageHeight =  1 

# Creating the folder where we dump the stencils
path = ".\mystencils"

try:
    os.mkdir(path)
except OSError:
    print (f'Creation of the directory {path} failed, or it already exists.')
else:
    print(f'Successfully created the directory {path}')


columns = [
	'Hostname',
	'Management Address',
	'Port',
	'Code',
	'Note',
	'Device Type',
	'Unique ID',
]

df = pd.DataFrame(columns = columns)

# Create the standardized .csv file and a copy:
df.to_csv('.\standardized.csv')

# Now the script needs to update the Visio document based on .csv input:

# 1) Create the Visio document next.

# Open Visio
visio = win32com.client.Dispatch('Visio.Application')

# Initialize a template, "Blank Drawing".
customTemplate = ''
docTemplate = visio.Documents.Add(customTemplate)

# Import the stencil (Visio document) containing all of the shapes
# required to run the script.
newStencil = 'C:\\Users\\Queen\\Documents\\My Shapes\\testStencil.vssx'
newStencilTemplate = visio.Documents.Add(newStencil)

# Use the first page, "Page-1", of the document.
pg = docTemplate.Pages.Item(1)

# Place a Visio shape on the Visio document
def dropShape (shapeType, posX, posY, theText):

    print(f"Shape type = {shapeType}")
    print(f"X = {posX}")
    print(f"Y = {posY}")

    vsoShape = pg.Drop(shapeType, posX, posY)
    vsoShape.Text = theText

    return vsoShape   # Returns the shape that was created

# Draw connector from bottom of one shape to another shape with autoroute
def connectShapes(shape1, shape2, theText):

    # Create and drop a dynamic Connector onto the page
    conn = visio.Application.ConnectorToolDataObject
    shpConn = pg.Drop(conn, 0, 0)

    shpConn.CellsU("BeginX").GlueTo(shape1.CellsU("PinX"))          
    shpConn.CellsU("EndX").GlueTo(shape2.CellsU("PinX"))

    shpConn.Text = theText

# Get the stencil object
def getStencilName(): # Name of Visio stencil containing shapes

    testStencilName = "Stencil2" # The custom stencil
    docFlowStencil = ""

    for doc in visio.Documents:
        print(f'Doc name = {doc}')
        if doc.Name == testStencilName or doc.Name == '' :

            docFlowStencil = doc

    print (f'docFlowStencil = {docFlowStencil}') # Print installed stencils
    return docFlowStencil

def main(x):

	# Output some data onto the console.
    docFlowStencil = getStencilName()

	# Name all of the shapes in your stencil. So far, I have included 2, but you may have many more shapes and that's okay.
    namesOfShapes = ['pc', 'ap']
    numberOfShapes = len(namesOfShapes)

    print(f'Shapes read: {numberOfShapes}')

    # Get masters for the devices, i.e. shapes:
    masterList = [docFlowStencil.Masters.ItemU(namesOfShapes[i]) for i in range(numberOfShapes)]

    # Setting co-ordinates...
    x = 7
    y = 8.5

    # Update the DataFrame when you have entered all of your details into the .csv:
    input('Press \'Enter\' if the standardized .csv file (named \'standardized.csv\' is updated AND saved:')
    df = pd.read_csv('.\standardized.csv')

	# Make a copy of the .csv.
	# REQUIRED because if the script is closed and re-run, the original .csv will be over-written.
    df.to_csv('.\standardized_copy.csv')

    print(df)
    
    numberOfDevices = df.shape[0]
    print(f'Devices read: {numberOfDevices}')

    # Time to drop the shapes from the master list onto the document Drawing1!
    # ds is short for droppedShapes
    ds = []
    j = 0
    for i in range(numberOfDevices):
    	if i != 0 and i % 4 == 0:
    		x = x + 2
    		j = 0
    	ds.append(dropShape(masterList[0], x, y + 1.5*j, df.at[i, 'Hostname']))
    	j += 1

    # Resizing the devices by 10 page units (code 63)
    for i in range(numberOfDevices):
    	ds[i].Resize(1, -10, 63)
  
    # Add connectors to the shapes based on the unique ID
    for i in range(numberOfDevices):
    	uID_i = df.at[i, 'Unique ID']
    	for j in range(numberOfDevices):
    		uID_j = df.at[j, 'Unique ID']
    		if i < j and uID_i == uID_j:
    			connectShapes(ds[i], ds[j], 'yourTextHere')
main(1)

# 2) Check for changes to the Excel spreadsheet: If it's changed (perhaps by 'save'),
#	 update the Visio file.