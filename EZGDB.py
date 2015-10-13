from openpyxl import load_workbook
from xml.etree.ElementTree import ElementTree as ET, XML
from xml.etree.ElementTree import ElementTree,Element, SubElement, Comment, tostring 
import arcpy
import collections

class Expando(object):
    pass

inworkbook=arcpy.GetParameterAsText(0)
inxmlfile=arcpy.GetParameterAsText(1)

wb='C://Users//sberg//OneDrive - Vanasse Hangen Brustlin, Inc-//Shared with Everyone//Runoff Tracking and Accounting//VHB Runoff Tracking and Accounting Architecture - Project.xlsx'
if inworkbook!=None and inworkbook!='': 
    arcpy.AddMessage("Input Workbook: " + inworkbook)
    wb=inworkbook

#f=open('C://data//_code//GDBModel//EZGDB//work.xml','w')
sOutXMLFile='C://Users//sberg//OneDrive - Vanasse Hangen Brustlin, Inc-//Shared with Everyone//Runoff Tracking and Accounting//Project.xml'# 'C://data//_code//GDBModel//EZGDB//work.xml'
sFCName="new"
eGT = 'esriGeometryPoint' #default
sWKT='GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137.0,298.257223563]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433],AUTHORITY["EPSG",4326]]'


if inxmlfile!=None and inxmlfile!='':sOutXMLFile=inxmlfile
arcpy.AddMessage("Output XML File: " + sOutXMLFile)

wb2 = load_workbook(wb)#'C://Users//sberg//OneDrive - Vanasse Hangen Brustlin, Inc-//Shared with Everyone//Runoff Tracking and Accounting//VHB Runoff Tracking and Accounting Architecture - Project.xlsx')#('C://data//_code//GDBModel//EZGDB//work.xlsx')


print 'reading sheets... ' + str(wb2.get_sheet_names())
#arcpy.AddMessage('reading sheets... ' + str(wb2.get_sheet_names()))

#read meta sheet
nme = wb2['Meta']['B1'].value
tpe=wb2['Meta']['B2'].value
eGT=tpe

#read choices sheet
choicesheet = wb2['choices']
highrow=choicesheet.get_highest_row()
domainlist= []
domainvalues=[]
rawdomainlist=[]

#for each domain value, need to group

for r in range(2,highrow+1):

    domainname=choicesheet['A' + str(r)].value
    domval=choicesheet['B' + str(r)].value
    domlbl=choicesheet['C' + str(r)].value
    d=None

    if(domainname!=None and domainname!='' ):
        #find domain name in all domainlist objects
        bDomainExists=False
        for d2 in domainlist:
            if (str(domainname).strip() == str(d2.domainname).strip()):
                d=d2
                break

        if(d==None):
            d=Expando()
            d.domainname = str(domainname).strip()
            d.values=[]
            domainlist.append(d)
        
        d.values.append({"domainvalue":domval,"domainlabel":domlbl})


dl = domainlist




#read fields sheet
fieldssheet=wb2['survey']
highrow=fieldssheet.get_highest_row()
fieldlist=[]
for row in range(2, highrow + 1):
    name  = fieldssheet['B' + str(row)].value
    type = fieldssheet['A' + str(row)].value
    label = fieldssheet['C' + str(row)].value
    fieldlist.append({"name":name,'type':type,'label':label})


print 'writing workspace to ' + sOutXMLFile + "..."
tree=ElementTree()

#ws=Element('esri:Workspace',{'xmlns:esri':'http://www.esri.com/schemas/ArcGIS/10.3','xmlns:xsi':'http://www.w3.org/2001/XMLSchema-instance','xmlns:xs':'http://www.w3.org/2001/XMLSchema'})
ws=Element('esri:Workspace',{'xmlns:esri':'http://www.esri.com/schemas/ArcGIS/10.3','xmlns:xs':'http://www.w3.org/2001/XMLSchema'})

wd=SubElement(ws,'WorkspaceDefinition',{'xsi:type':'esri:WorkspaceDefinition'})

wt=SubElement(wd,'WorkspaceType').text="esriLocalDatabaseWorkspace"

v=SubElement(wd,'Version')

domains=SubElement(wd,'Domains',{'xsi:type':'esri:ArrayOfDomain'})

#add domains
for domaintoadd in dl:
    dmaine=SubElement(domains,'Domain',{'xsi:type':'esri:CodedValueDomain'})
    SubElement(dmaine,"DomainName").text = domaintoadd.domainname
    SubElement(dmaine,"FieldType").text = 'esriFieldTypeString'
    SubElement(dmaine,"MergePolicy").text = 'esriMPTDefaultValue'
    SubElement(dmaine,"SplitPolicy").text = 'esriSPTDefaultValue'
    SubElement(dmaine,"Description").text = ''
    SubElement(dmaine,"Owner").text = ''
    cvals = SubElement(dmaine,"CodedValues",{'xsi:type':'esri:ArrayOfCodedValue'})
    for cv in domaintoadd.values:
        cv2= SubElement(cvals,'CodedValue',{'xsi:type':'esri:CodedValue'})
        SubElement(cv2,"Name").text = cv['domainlabel']
        SubElement(cv2,"Code").text = cv['domainlabel']
    domaintoadd.element = dmaine
     

dd=SubElement(wd,'DatasetDefinitions',{'xsi:type':'esri:ArrayOfDataElement'})
de=SubElement(dd,'DataElement',{'xsi:type':'esri:DEFeatureClass'})
cp=SubElement(de,'CatalogPath').text="/FC=" + sFCName
name=SubElement(de,'Name').text= nme# sFCName
dt=SubElement(de,'DatasetType').text='esriDTFeatureClass'
dsid=SubElement(de,'DSID').text= "0"
hid=SubElement(de,'HasOID').text='true'
oid=SubElement(de,'OIDFieldName').text='OBJECTID'
fields=SubElement(de,'Fields',{'xsi:type':'esri:Fields'})
fieldarray=SubElement(fields,'FieldArray',{'xsi:type':'esri:ArrayOfField'})

field1=SubElement(fieldarray,'Field',{'xsi:type':'esri:Field'})
fieldname=SubElement(field1,'Name').text="OBJECTID"
fieldtype=SubElement(field1,'Type').text="esriFieldTypeOID"
fieldisnull=SubElement(field1,'IsNullable').text="false"
fieldlength=SubElement(field1,'Length').text="4"
fieldprecision=SubElement(field1,'Precision').text="0"
fieldscale=SubElement(field1,'Scale').text="0"
fieldrequired=SubElement(field1,'Required').text="true"
fieldeditable=SubElement(field1,'Editable').text="false"
fielddomainfixed=SubElement(field1,'DomainFixed').text="true"
fieldaliasname=SubElement(field1,'AliasName').text="OBJECTID"
fieldmodelname=SubElement(field1,'ModelName').text="OBJECTID"

field2=SubElement(fieldarray,'Field',{'xsi:type':'esri:Field'})
fieldname=SubElement(field2,'Name').text="SHAPE"
fieldtype=SubElement(field2,'Type').text="esriFieldTypeGeometry"
fieldisnull=SubElement(field2,'IsNullable').text="true"
fieldlength=SubElement(field2,'Length').text="0"
fieldprecision=SubElement(field2,'Precision').text="0"
fieldscale=SubElement(field2,'Scale').text="0"
fieldrequired=SubElement(field2,'Required').text="true"
fielddomainfixed=SubElement(field2,'DomainFixed').text="true"
fieldaliasname=SubElement(field2,'AliasName').text="SHAPE"
fieldmodelname=SubElement(field2,'ModelName').text="SHAPE"


#add specified fields
for fld in fieldlist:
    if fld['name']!=None and fld['name']!='':
        newfield=SubElement(fieldarray,'Field',{'xsi:type':'esri:Field'})
        fieldname=SubElement(newfield,'Name').text=str(fld['name']).strip()
        t=fld['type']
        ftpe="esriFieldTypeString"
        doma=""
        if(t.startswith('text')): ftpe="esriFieldTypeString"
        if(t.startswith('integer')): ftpe="esriFieldTypeInteger"
        if(t.startswith('decimal')): ftpe="esriFieldTypeDouble"
        if(t.startswith('date')): ftpe="esriFieldTypeDate"
        if(t.startswith('select_one')): 
            ftpe="esriFieldTypeString"
            doma=t.split()[1]
            doma=doma.strip()

        fieldtype=SubElement(newfield,'Type').text=ftpe#"esriFieldTypeString" #todo fld['type'] via lookup
        fieldisnull=SubElement(newfield,'IsNullable').text="true"
        fieldlength=SubElement(newfield,'Length').text="255"
        fieldprecision=SubElement(newfield,'Precision').text="0"
        fieldscale=SubElement(newfield,'Scale').text="0"
        fieldrequired=SubElement(newfield,'Required').text="true"
        fielddomainfixed=SubElement(newfield,'DomainFixed').text="true"
        fieldaliasname=SubElement(newfield,'AliasName').text=fld['label']
        fieldmodelname=SubElement(newfield,'ModelName').text=fld['name']

        if doma!='':
        #dmaine=SubElement(newfield,'Domain',{'xsi:type':'esri:CodedValueDomain'})
        #find domain element
            for domaintoadd in dl:
                 if domaintoadd.domainname==doma:
                    newfield.append(domaintoadd.element)
                    break
                

         
####

geomdef=SubElement(field2,'GeometryDef',{'xsi:type':'esri:GeometryDef'})
a=SubElement(geomdef,'AvgNumPoints').text='0'
gt=SubElement(geomdef,'GeometryType').text=eGT
hm=SubElement(geomdef,'HasM').text='false'
hz=SubElement(geomdef,'HasZ').text='false'
sr=SubElement(geomdef,'SpatialReference',{'xsi:type':'esri:GeographicCoordinateSystem'})
a1=SubElement(sr,'XOrigin').text='-400'
a2=SubElement(sr,'YOrigin').text='-400'
a3=SubElement(sr,'XYScale').text='999999999.99999988'
a4=SubElement(sr,'ZOrigin').text='-100000'
a5=SubElement(sr,'MOrigin').text='10000'
a6=SubElement(sr,'MScale').text='-10000'
a7=SubElement(sr,'XYTolerance').text='8.983152841195215e-009'
a8=SubElement(sr,'ZTolerance').text='0.001'
a9=SubElement(sr,'MTolerance').text='0.001'
a10=SubElement(sr,'HighPrecision').text='true'
a11=SubElement(sr,'LeftLongitude').text='-180'
a12=SubElement(sr,'WKID').text='4326'
a13=SubElement(sr,'LatestWKID').text='4326'

gs=SubElement(geomdef,'GridSize0').text='0'

indexes=XML('<Indexes xsi:type="esri:Indexes" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ><IndexArray xsi:type="esri:ArrayOfIndex"><Index xsi:type="esri:Index"><Name>FDO_OBJECTID</Name><IsUnique>true</IsUnique><IsAscending>true</IsAscending><Fields xsi:type="esri:Fields"><FieldArray xsi:type="esri:ArrayOfField"><Field xsi:type="esri:Field"><Name>OBJECTID</Name><Type>esriFieldTypeOID</Type><IsNullable>false</IsNullable><Length>4</Length><Precision>0</Precision><Scale>0</Scale><Required>true</Required><Editable>false</Editable><DomainFixed>true</DomainFixed><AliasName>OBJECTID</AliasName><ModelName>OBJECTID</ModelName></Field></FieldArray></Fields></Index><Index xsi:type="esri:Index"><Name>FDO_SHAPE</Name><IsUnique>false</IsUnique><IsAscending>true</IsAscending><Fields xsi:type="esri:Fields"><FieldArray xsi:type="esri:ArrayOfField"><Field xsi:type="esri:Field"><Name>SHAPE</Name><Type>esriFieldTypeGeometry</Type><IsNullable>true</IsNullable><Length>0</Length><Precision>0</Precision><Scale>0</Scale><Required>true</Required><DomainFixed>true</DomainFixed><GeometryDef xsi:type="esri:GeometryDef"><AvgNumPoints>0</AvgNumPoints><GeometryType>' + eGT + '</GeometryType><HasM>false</HasM><HasZ>false</HasZ><SpatialReference xsi:type="esri:ProjectedCoordinateSystem"><WKT>PROJCS["NAD_1983_StatePlane_Massachusetts_Mainland_FIPS_2001",GEOGCS["GCS_North_American_1983",DATUM["D_North_American_1983",SPHEROID["GRS_1980",6378137.0,298.257222101]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433]],PROJECTION["Lambert_Conformal_Conic"],PARAMETER["False_Easting",200000.0],PARAMETER["False_Northing",750000.0],PARAMETER["Central_Meridian",-71.5],PARAMETER["Standard_Parallel_1",41.71666666666667],PARAMETER["Standard_Parallel_2",42.68333333333333],PARAMETER["Latitude_Of_Origin",41.0],UNIT["Meter",1.0],AUTHORITY["EPSG",26986]]</WKT><XOrigin>-36530900</XOrigin><YOrigin>-28803200</YOrigin><XYScale>10000</XYScale><ZOrigin>-100000</ZOrigin><ZScale>10000</ZScale><MOrigin>-100000</MOrigin><MScale>10000</MScale><XYTolerance>0.001</XYTolerance><ZTolerance>0.001</ZTolerance><MTolerance>0.001</MTolerance><HighPrecision>true</HighPrecision><WKID>26986</WKID><LatestWKID>26986</LatestWKID></SpatialReference><GridSize0>12000</GridSize0></GeometryDef><AliasName>SHAPE</AliasName><ModelName>SHAPE</ModelName></Field></FieldArray></Fields></Index></IndexArray></Indexes>')
de.append(indexes)
ep=XML('<ExtensionProperties xsi:type="esri:PropertySet" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><PropertyArray xsi:type="esri:ArrayOfPropertySetProperty"></PropertyArray></ExtensionProperties>')
de.append(ep)

ft=SubElement(de,'FeatureType').text='esriFTSimple'
ft2=SubElement(de,'ShapeType').text=eGT
ft3=SubElement(de,'ShapeFieldName').text='SHAPE'
ft4=SubElement(de,'HasM').text='false'
ft5=SubElement(de,'HasZ').text='false'
ft6=SubElement(de,'HasSpatialIndex').text='true'
ft7=SubElement(de,'AreaFieldName').text=''
ft8=SubElement(de,'LengthFieldName').text=''

ext=XML('<Extent xsi:type="esri:EnvelopeN" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><XMin>NaN</XMin><YMin>NaN</YMin><XMax>NaN</XMax><YMax>NaN</YMax><SpatialReference xsi:type="esri:GeographicCoordinateSystem"><WKT>GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137.0,298.257223563]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433],AUTHORITY["EPSG",4326]]</WKT><XOrigin>-400</XOrigin><YOrigin>-400</YOrigin><XYScale>999999999.99999988</XYScale><ZOrigin>-100000</ZOrigin><ZScale>10000</ZScale><MOrigin>-100000</MOrigin><MScale>10000</MScale><XYTolerance>8.983152841195215e-009</XYTolerance><ZTolerance>0.001</ZTolerance><MTolerance>0.001</MTolerance><HighPrecision>true</HighPrecision><LeftLongitude>-180</LeftLongitude><WKID>4326</WKID><LatestWKID>4326</LatestWKID></SpatialReference></Extent>')
de.append(ext)

sr3=XML('<SpatialReference xsi:type="esri:GeographicCoordinateSystem" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><WKT>GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137.0,298.257223563]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433],AUTHORITY["EPSG",4326]]</WKT><XOrigin>-400</XOrigin><YOrigin>-400</YOrigin><XYScale>999999999.99999988</XYScale><ZOrigin>-100000</ZOrigin><ZScale>10000</ZScale><MOrigin>-100000</MOrigin><MScale>10000</MScale><XYTolerance>8.983152841195215e-009</XYTolerance><ZTolerance>0.001</ZTolerance><MTolerance>0.001</MTolerance><HighPrecision>true</HighPrecision><LeftLongitude>-180</LeftLongitude><WKID>4326</WKID><LatestWKID>4326</LatestWKID></SpatialReference>')
de.append(sr3)

ct=SubElement(de,'ChangeTracked').text='false'
wd=SubElement(ws,'WorkspaceData',{'xsi:type':'esriWorkspaceData'})

tree._setroot(ws)
tree.write(sOutXMLFile)

arcpy.SetParameter(1,sOutXMLFile)

print 'complete.'


#f.write(ws)
#f.close()
