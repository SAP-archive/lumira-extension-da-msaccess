#-----------------------------------------------------------------------------------------------|---------------imports#
import win32com.client
import easygui
import logging
import sys
import tempfile
import os
import pythoncom
import decimal

reload(sys)
sys.setdefaultencoding("utf-8")


#-----------------------------------------------------------------------------------------------|--variable-definitions#
def enum(*sequential, **named):
    enums = dict(zip(sequential, range(len(sequential))), **named)
    return type('Enum', (), enums)


Mode = enum('PREVIEW', 'EDIT', 'REFRESH')
mode = 0
catalog = win32com.client.Dispatch('ADOX.Catalog')
connection = win32com.client.Dispatch('ADODB.Connection')
connection.Cursorlocation = 3  #Use local cursor
command = win32com.client.Dispatch('ADODB.Command')
recordset = win32com.client.Dispatch('ADODB.Recordset')
parameters = win32com.client.Dispatch('ADODB.Parameter')
#field = win32com.client.Dispatch('ADODB.Field')
fieldNames = []
fieldValues = []
fileNameAndPath = ''
procedure_choice = ''
params = ''
paramslist = []

#--------------------------------------------------------------------------------------Parse-Command-Line-Parameters---#
for i in range(len(sys.argv)):
    if str(sys.argv[i]).lower() == "-mode" and (i + 1) < len(sys.argv):
        if str(sys.argv[i + 1]).lower() == "preview":
            mode = Mode.PREVIEW
        elif str(sys.argv[i + 1]).lower() == "edit":
            mode = Mode.EDIT
        elif str(sys.argv[i + 1]).lower() == "refresh":
            mode = Mode.REFRESH
    elif str(sys.argv[i]).lower() == "-size":
        size = int(sys.argv[i + 1])
    elif str(sys.argv[i]).lower() == "-params":
        params = str(sys.argv[i + 1])
        paramslist = params.split(';')
    i += 1
#easygui.msgbox(str(paramslist))
#----------------------------------------------------------------------------------------------------parseArgs---------#
def parseArgs():
    global fieldValues
    global fileNameAndPath
    global procedure_choice

    #if paramslist is None: break
    i = 0
    q = 0
    for i in range(len(paramslist)):
        if paramslist[i].split('=')[0].lower() == 'file_location':
            fileNameAndPath = paramslist[i].split('=')[1]
        elif paramslist[i].split('=')[0].lower() == 'procedure':
            procedure_choice = paramslist[i].split('=')[1]
        elif "parameter_" in paramslist[i].split('=')[0].lower():
            fieldNames.append(paramslist[i].split('=')[0].replace("parameter_", ""))
            fieldValues.append(paramslist[i].split('=')[1])
            q += 1
        i += 1


def printData(screenInput):
    global connection
    global catalog
    global recordset
    global fieldNames
    global fieldValues
    global fileNameAndPath
    global procedure_choice

    #-------------------------------------------------------------------------------------------|-------open-connection#
    if screenInput == 1:
        fileNameAndPath = easygui.fileopenbox(title="Open Access Database File", filetypes=["*.accdb"])
        if fileNameAndPath == None:
            sys.exit("User Cancelled")
    connectionString = ''.join(['Provider=Microsoft.ACE.OLEDB.12.0;Data Source=', fileNameAndPath])
    try:
        connection.Open(connectionString)
    except:
        sys.exit("Error Opening File, Wrong File Type?")
    catalog.ActiveConnection = connection
    recordset.ActiveConnection = connection

    #-------------------------------------------------------------------------------------------|-------query-selection#
    logging.info(''.join(['--procedures : ', str(catalog.procedures.count)]))
    catalog.procedures.Refresh()
    procedure_choices = []
    i = 0
    for i in range(catalog.procedures.count):
        procedure_choices.append(catalog.procedures[i].name)
    if screenInput == 1:
        procedure_choice = easygui.choicebox('Choose A Query From List', 'Query List', procedure_choices)
        if procedure_choice == None:
            sys.exit("User Cancelled")
    logging.info(''.join(['--procedure_choice : ', str(procedure_choice)]))


    #-------------------------------------------------------------------------------------------|-----set-parameters---#
    command = catalog.procedures(procedure_choice).command
    parameters = command.parameters

    if screenInput == 1:
        i = 0
        errmsg = ''
        #Prepare popup catalog
        fieldNamesbak = fieldNames
        fieldValuesbak = fieldValues
        fieldNames = []
        fieldValues = []

        #easygui.msgbox(''.join(['--', 'length of fieldValues', ' : ', str(len(fieldValues))]))
        for i in range(command.parameters.count):
            if str(command.parameters(i).name) not in fieldNamesbak:
                fieldNames.append(str(command.parameters(i).name))
                fieldValues.append('')
            else:
                fieldNames.append(str(command.parameters(i).name))
                fieldValues.append(fieldValuesbak[fieldNamesbak.index(str(command.parameters(i).name))])
            #Generate popup

        fieldValues = easygui.multenterbox(errmsg, 'Enter Parameter Values', fieldNames, fieldValues)
        while 1:
            if fieldValues == None: break
            errmsg = ""
            for i in range(len(fieldNames)):
                if fieldValues[i].strip() == "":
                    errmsg += ('"%s" is a required field.\n\n' % fieldNames[i])
            if errmsg == "":
                break  # no problems found
            fieldValues = easygui.multenterbox(errmsg, 'Enter Parameter Values', fieldNames, fieldValues)
            if fieldValues == None:
                sys.exit("User Cancelled")

    #set parameter values according to popup
    i = 0
    #command.parameters.refresh
    for i in range(len(fieldValues)):
        #easygui.msgbox(str(i))
        try:
            parameters(str(fieldNames[i])).Value = str(fieldValues[i])
            parameters.append
        except pythoncom.com_error:
            pass

    #-------------------------------------------------------------------------------------------|-----execute-query----#
    (recordset, result) = command.Execute()


    #-------------------------------------------------------------------------------------------|-----print-DS-Info----#
    print "beginDSInfo"
    print "csv_first_row_has_column_names;true;true"
    print "csv_separator;,;true"
    print "csv_number_grouping;,;true"
    print "csv_number_decimal;.;true"
    print "csv_date_format;d.M.yyyy;true"
    print ''.join(['file_location;', fileNameAndPath, ';true'])
    print ''.join(['procedure;', procedure_choice, ';true'])
    i = 0
    #easygui.msgbox(''.join(['--', 'length of fieldValues', ' : ', str(len(fieldValues))]))
    #easygui.msgbox(''.join(['--', 'length of fieldNames', ' : ', str(fieldNames)]))
    for i in range(len(fieldNames)):
        print ''.join(['parameter_', fieldNames[i], ';', fieldValues[i], ';true'])
    print "endDSInfo"
    print "beginData"




    #-------------------------------------------------------------------------------------------|-----write-results----#
    if recordset.RecordCount == 0:
        print "Error"
        print "No Records"
        print "endData"
    else:
        for i in range(len(recordset.Fields)):
            if i > 0:
                sys.stdout.softspace=0
                print ', ',
            sys.stdout.softspace=0
            print str(recordset.Fields(i).name),
        sys.stdout.softspace=0
        print ''
        while not recordset.EOF:
            for i in range(len(recordset.Fields)):
                if i > 0:
                    sys.stdout.softspace=0
                    print ', ',
                sys.stdout.softspace=0
                print str(recordset.Fields(i).value),
            sys.stdout.softspace=0
            print ''
            recordset.MoveNext()
        recordset.Close()
        print "endData"
    connection.Close
    del catalog
    del connection
    del command
    del recordset
    del parameters

#----------------------------------------------------------------------------------------------------Mode-Logic--------#

LOG_FILENAME = tempfile.NamedTemporaryFile(delete=False)
logging.basicConfig(filename=LOG_FILENAME.name,level=logging.DEBUG,)

try:
    if mode == Mode.PREVIEW:
        printData(screenInput = 1)
    elif mode == Mode.EDIT:
        parseArgs()
        printData(screenInput = 1)
    elif mode == Mode.REFRESH:
        parseArgs()
        printData(screenInput = 0)
except SystemExit as e:
    print "beginData"
    print "Error"
    print e
    print "endData"
except:
   logging.exception('Got exception on main handler')
   os.system("notepad " + LOG_FILENAME.name)
   #raise

