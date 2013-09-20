#-------------------------------------------------------------
# Name:       Desktop Word Report
# Purpose:    This script will produce a PDF with two pages - One for a report and one for a map. Input to this script would
#             be a property feature class and a feature class that is to be reported on. Multiple properties to one mail address 
#             is accounted for, so one report will be produced in this case with many maps (depending on how many properties).        
# Author:     Shaun Weston (shaun.weston@splicegroup.co.nz)
# Date Created:    09/11/2011
# Last Updated:    20/09/2013
# Copyright:   (c) Splice Group
# ArcGIS Version:   10.1/10.2
# Python Version:   2.7
#--------------------------------

# Import modules and enable data to be overwritten
import os
import sys
import string
import datetime
import smtplib
import zipfile
import shutil
import win32com.client
import arcpy
arcpy.env.overwriteOutput = True

# Set variables
logInfo = "false"
logFile = r""
sendEmail = "false"
output = None

# Start of main function
def mainFunction(propertyFeatureClass,analysisFeatureClass,propertyIDField,mailAddressField,reportFields,reportFieldPlaceholders,mxdTemplate,layerSymbology,wordTemplate,outputFolder): # Get parameters from ArcGIS Desktop tool by seperating by comma e.g. (var1 is 1st parameter,var2 is 2nd parameter,var3 is 3rd parameter)  
    try:
        # Log start
        if logInfo == "true":
            loggingFunction(logFile,"start","")

        # --------------------------------------- Start of code --------------------------------------- #        

        # Spatially join analysis feature class to property        
        arcpy.SpatialJoin_analysis(propertyFeatureClass, analysisFeatureClass, "in_memory\PropertyAffected", "JOIN_ONE_TO_MANY", "KEEP_COMMON", "#", "INTERSECT")

        # Get the number of duplicate mail addresses
        arcpy.Frequency_analysis("in_memory\PropertyAffected", "in_memory\PropertyAffectedFrequency", mailAddressField, "")        
        arcpy.JoinField_management("in_memory\PropertyAffected", mailAddressField, "in_memory\PropertyAffectedFrequency", mailAddressField, "FREQUENCY")

        # Add on report added field
        arcpy.AddField_management("in_memory\PropertyAffected", "ReportAdded", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
        arcpy.CalculateField_management("in_memory\PropertyAffected", "ReportAdded", "\"No\"", "PYTHON_9.3", "")

        # Add on another valuation field - For adding maps to same mail address
        arcpy.AddField_management("in_memory\PropertyAffected", "FirstProperty", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
        arcpy.CalculateField_management("in_memory\PropertyAffected", "FirstProperty", "!" + propertyIDField + "!", "PYTHON_9.3", "")

        # Setup map document
        mxd = arcpy.mapping.MapDocument(mxdTemplate)        
        # Reference data frame and the layer
        dataFrame = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
        
        # Add the affected properties to the map
        arcpy.MakeFeatureLayer_management("in_memory\PropertyAffected", "Properties Affected")  
        layer = arcpy.mapping.Layer("Properties Affected")
        arcpy.mapping.AddLayer(dataFrame,layer)
        layer = arcpy.mapping.ListLayers(mxd, "Properties Affected", dataFrame)[0]
        # Update the symbology
        symbologyLayer = arcpy.mapping.Layer(layerSymbology)
        arcpy.mapping.UpdateLayer(dataFrame, layer, symbologyLayer, True)

        # Setup outpath paths
        zipOutputPath = outputFolder + '\\Zip\\'
        docOutputPath = outputFolder + '\\Docs\\'
        # If it doesn't exist, create it
        if not os.path.exists(docOutputPath):
            os.makedirs(docOutputPath)

        # Setup the feature class to read through each of the attributes
        rows = arcpy.SearchCursor("in_memory\PropertyAffected")
        row = rows.next()

        # Do the following until all the rows have been read       
        while row:
                # Define the attributes required for the report       
                OBJECTID = str(row.OBJECTID)
                NumProperties = row.FREQUENCY
                ReportAdded = str(row.ReportAdded)
                FirstProperty = str(row.FirstProperty)
                PropertyID = str(row.getValue(propertyIDField))
                Owner_Mail = str(row.getValue(mailAddressField))
  
                # This will check to see if a report has already been created for this mail address
                # If yes then only produce a map for the next property
                if ReportAdded == 'Yes':

                    # Open up the map document data frame, select the feature class and zoom to it
                    layer = arcpy.mapping.ListLayers(mxd, "Properties Affected", dataFrame)[0]
                    arcpy.SelectLayerByAttribute_management(layer, "NEW_SELECTION", '"OBJECTID" =' + str(OBJECTID))
                    dataFrame.extent = layer.getSelectedExtent(False)
                    trueScale = dataFrame.scale * 4
                    #Round scale to a more general number
                    dataFrame.scale = round(trueScale, -2)
                    arcpy.RefreshActiveView()
                
                    # Export to PDF
                    arcpy.mapping.ExportToPDF(mxd, docOutputPath + "\\Map" + PropertyID + ".pdf")           
                # If no produce a report and map
                else:
                    # Unzip the word document to Open XML format and assign the document xml (contains the main content) to a variable
                    zipDoc = zipfile.ZipFile(wordTemplate)
                    zipDoc.extractall(zipOutputPath)
                    wordDocXML = zipOutputPath + '\\word\\document.xml'
                
                    # A find and replace  on the word document XML file
                    s = open(wordDocXML).read()
                    
                    # Replace mail address
                    s = s.replace('XMAILADDRESSX', Owner_Mail)

                    # Put required fields and placeholders into an array
                    # If a string, convert to array
                    if isinstance(reportFields, basestring):
                        reportFields = string.split(reportFields, ";")
                    if isinstance(reportFieldPlaceholders, basestring):                   
                        reportFieldPlaceholders = string.split(reportFieldPlaceholders, ";")

                    # Loop through each of the report fields
                    count = 0
                    while (len(reportFields) > count):
                        # Find and replace text
                        s = s.replace(str(reportFieldPlaceholders[count]), str(row.getValue(reportFields[count])))
                        count = count + 1

                    f = open(wordDocXML, 'w')
                    f.write(s)              
                    f.close()

                    newDocZip = zipfile.ZipFile(docOutputPath + "\\Report" + PropertyID + ".docx", "w")
                    root_len = len(os.path.abspath(zipOutputPath))
                    for root, dirs, files in os.walk(zipOutputPath):
                        archive_root = os.path.abspath(root)[root_len:]
                        for f in files:
                            fullpath = os.path.join(root, f)
                            archive_name = os.path.join(archive_root, f)
                            newDocZip.write(fullpath, archive_name)
                    newDocZip.close()

                    # Open up word document                    
                    app = win32com.client.Dispatch('Word.Application')
                    app.Visible = 0
                    app.DisplayAlerts = 0                    
                    app.Documents.Open(docOutputPath + "\\Report" + PropertyID + ".docx")                   
                    doc = app.ActiveDocument
                    # Save as PDF
                    doc.SaveAs(docOutputPath + "\\Report" + PropertyID + ".pdf", FileFormat=17)                                      
                    doc.Close
                    app.Quit()

                    # Open up the map document data frame, select the feature class and zoom to it
                    layer = arcpy.mapping.ListLayers(mxd, "Properties Affected", dataFrame)[0]
                    arcpy.SelectLayerByAttribute_management(layer, "NEW_SELECTION", '"OBJECTID" =' + str(OBJECTID))
                    dataFrame.extent = layer.getSelectedExtent(False)
                    trueScale = dataFrame.scale * 4
                    # Round scale to a more general number
                    dataFrame.scale = round(trueScale, -2)
                    arcpy.RefreshActiveView()
                
                    # Export to PDF
                    arcpy.mapping.ExportToPDF(mxd, docOutputPath + "\\Map" + PropertyID + ".pdf")

                # This section will append the report and map into a final report, first we need to see if it has already been created for
                # mail address
                # If yes, we need to open that pdf up and attach on another map to it.
                pdfPath = outputFolder + "\\ReportWithMap" + FirstProperty + ".pdf"    
                if os.path.exists(pdfPath):
                    pdfReport = arcpy.mapping.PDFDocumentOpen(outputFolder + "\\ReportWithMap" + FirstProperty + ".pdf")
                    if ReportAdded == 'No':
                        pdfReport.appendPages(docOutputPath + "\\Report" + PropertyID + ".pdf")
                        pdfReport.appendPages(docOutputPath + "\\Map" + PropertyID + ".pdf")
                        pdfReport.saveAndClose()

                        if NumProperties > 1:
                            query = "if [Owner_Mail] = " + "\"" + Owner_Mail + "\" " + "Then\\n    var = \"Yes\"\\nElse\\n    var = [ReportAdded]\\nEnd If"
                            query2 = "if [Owner_Mail] = " + "\"" + Owner_Mail + "\" " + "Then\\n    var2 = " + "\"" + PropertyID + "\" " + "\\nElse \\n    var2 = [FirstProperty] \\nEnd If"
                            arcpy.CalculateField_management("in_memory\PropertyAffected", "ReportAdded", "var", "VB", query)
                            arcpy.CalculateField_management("in_memory\PropertyAffected", "FirstProperty", "var2", "VB", query2)
                        
                        # Go to the next row until the while loop has finished        
                        row = rows.next()           
                    else:           
                        pdfReport.appendPages(docOutputPath + "\\Map" + PropertyID + ".pdf")            
                        pdfReport.saveAndClose()
                        # Go to the next row until the while loop has finished        
                        row = rows.next()
                # If no, we create a new pdf and append the pages together.      
                else:    
                    pdfReport = arcpy.mapping.PDFDocumentCreate(outputFolder + "\\ReportWithMap" + FirstProperty + ".pdf")
                    if ReportAdded == 'No':
                        pdfReport.appendPages(docOutputPath + "\\Report" + PropertyID + ".pdf")
                        pdfReport.appendPages(docOutputPath + "\\Map" + PropertyID + ".pdf")
                        pdfReport.saveAndClose()

                        if NumProperties > 1:
                            query = "if [Owner_Mail] = " + "\"" + Owner_Mail + "\" " + "Then\\n    var = \"Yes\"\\nElse\\n    var = [ReportAdded]\\nEnd If"
                            query2 = "if [Owner_Mail] = " + "\"" + Owner_Mail + "\" " + "Then\\n    var2 = " + "\"" + PropertyID + "\" " + "\\nElse \\n    var2 = [FirstProperty] \\nEnd If"
                            arcpy.CalculateField_management("in_memory\PropertyAffected", "ReportAdded", "var", "VB", query)
                            arcpy.CalculateField_management("in_memory\PropertyAffected", "FirstProperty", "var2", "VB", query2)

                        # Go to the next row until the while loop has finished        
                        row = rows.next()           
                    else:           
                        pdfReport.appendPages(docOutputPath + "\\Map" + PropertyID + ".pdf")            
                        pdfReport.saveAndClose()
                        # Go to the next row until the while loop has finished        
                        row = rows.next()

        # Remove temporary folders                 
        shutil.rmtree(zipOutputPath)
        shutil.rmtree(docOutputPath)

        # --------------------------------------- End of code --------------------------------------- #  
            
        # If called from gp tool return the arcpy parameter   
        if __name__ == '__main__':
            # Return the output if there is any
            if output:
                arcpy.SetParameterAsText(1, output)
        # Otherwise return the result          
        else:
            # Return the output if there is any
            if output:
                return output      
        # Log start
        if logInfo == "true":
            loggingFunction(logFile,"end","")        
        pass
    # If arcpy error
    except arcpy.ExecuteError:
        # Show the message
        arcpy.AddMessage(arcpy.GetMessages(2))
        # Log error
        if logInfo == "true":  
            loggingFunction(logFile,"error",arcpy.GetMessages(2))
    # If python error
    except Exception as e:
        # Show the message
        arcpy.AddMessage(e.args[0])
        # Log error
        if logInfo == "true":         
            loggingFunction(logFile,"error",e.args[0])
# End of main function

# Start of logging function
def loggingFunction(logFile,result,info):
    #Get the time/date
    setDateTime = datetime.datetime.now()
    currentDateTime = setDateTime.strftime("%d/%m/%Y - %H:%M:%S")
    
    # Open log file to log message and time/date
    if result == "start":
        with open(logFile, "a") as f:
            f.write("---" + "\n" + "Process started at " + currentDateTime)
    if result == "end":
        with open(logFile, "a") as f:
            f.write("\n" + "Process ended at " + currentDateTime + "\n")
            f.write("---" + "\n")        
    if result == "error":
        with open(logFile, "a") as f:
            f.write("\n" + "Process ended at " + currentDateTime + "\n")
            f.write("There was an error: " + info + "\n")        
            f.write("---" + "\n")
        # Send an email
        if sendEmail == "true":
            arcpy.AddMessage("Sending email...")
            # Receiver email address
            to = ''
            # Sender email address and password
            gmail_user = ''
            gmail_pwd = ''
            # Server and port information
            smtpserver = smtplib.SMTP("smtp.gmail.com",587) 
            smtpserver.ehlo()
            smtpserver.starttls() 
            smtpserver.ehlo
            # Login
            smtpserver.login(gmail_user, gmail_pwd)
            # Email content
            header = 'To:' + to + '\n' + 'From: ' + gmail_user + '\n' + 'Subject:Error \n'
            msg = header + '\n' + '' + '\n' + '\n' + info
            # Send the email and close the connection
            smtpserver.sendmail(gmail_user, to, msg)
            smtpserver.close()                
# End of logging function    

# This test allows the script to be used from the operating
# system command prompt (stand-alone), in a Python IDE, 
# as a geoprocessing script tool, or as a module imported in
# another script
if __name__ == '__main__':
    # Arguments are optional - If running from ArcGIS Desktop tool, parameters will be loaded into *argv
    argv = tuple(arcpy.GetParameterAsText(i)
        for i in range(arcpy.GetArgumentCount()))
    mainFunction(*argv)
    
