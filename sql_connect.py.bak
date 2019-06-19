import pyodbc 
# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'VOO1DB2.ILPROD.LOCAL' 
database = 'ResearchMarketing' 
username = 'IMAGINELEARNING\kelly.richardson' 
password = 'grace7772' 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
print cursor