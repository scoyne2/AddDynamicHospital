' Author: coynes
' Change history:  V 1.0 coynes 6/4/2015  Created File
' Description: This file will dynamically add a Hospital column and populated the column based upon the  filename.
'              search for +++ to find the section that you should edit
'
' Inputs 
'   Filepath
'
' Outputs
'   File containing populated hospital column

'______________________________________________________________________________________ Setup Variables
Option Explicit
Dim objConcatedName, objDataExtension
Dim objDataFilename, objDataFullFileName
Dim objDataFSOin, objDataFilein, objFSOout, objFileout
Dim strRow, strHeader, strSpaces, strSeparator, strHospitalCode
Dim strFileNm
Dim i

Const ForReading = 1
Const ForWriting = 2

i=0
strRow = ""
strSpaces = "     "
objDataFilename = strSpaces	

'______________________________________________________________________________________ Capture and Evaluate Input
objDataFilename = Wscript.Arguments(0)
objDataFullFileName = objDataFileName

'______________________________________________________________________________________ Build File
'Use data file path and name as base for concatenated file name
If len(objDataFilename)>4 then
     If objDataFilename=strSpaces then '..do nothing
        else
          objDataExtension = right(ObjDataFullFileName,4)
          objConcatedName = left(ObjDataFullFileName,len(objDataFullFileName)-4) & "_cleaned" & objDataExtension
       end If 
   else
       objConcatedName = ObjDataFullFileName & "_concatenated"
end If

'______________________________________________________________________________________ Setup File
Set objDataFSOin = CreateObject("Scripting.FileSystemObject")
Set objFSOout = CreateObject("Scripting.FileSystemObject")

Set objDataFilein = objDataFSOin.OpenTextFile(objDataFullFileName, ForReading)
Set objFileout = objFSOout.OpenTextFile(objConcatedName, ForWriting,True)
      
'______________________________________________________________________________________ Process Data
'determine the hospitalcode based upon the file name
strFileNm = objDataFSOin.GetFileName(objDataFullFileName)
'write to console for testing
Wscript.Echo strFileNm 

'+++ +++ +++ +++ +++ +++ +++ +++ +++ +++ +++ +++ EDIT THIS CODE
'choose a delimiter
strSeparator = ","  

'create the header row
if i = 0 then    
    'if this is the first row of the file, add the column 'Hospital'
    strRow = objDataFilein.ReadLine
    strHeader = strRow & strSeparator & "Hospital"
Else     
  'if this is not the first row of the file, something went wrong
    Wscript.Echo( "Something went wrong, it looks like your incremental variable i did not start at 0" ) 
End if


Select Case strFileNm
  Case "SHECCorporateRoster.txt"
        'write to console for testing
        Wscript.Echo( "SHE" )   
        strHospitalCode = "SHE"
  Case "SJHCorporateRoster"
        'write to console for testing 
        Wscript.Echo( "SJH" )   
        strHospitalCode = "SJH"
  Case "SJBCorporateRoster"
        'write to console for testing
        Wscript.Echo( "SJB")   
        strHospitalCode = "SJB"
  'dont change this, there must be an ELSE statement incase no match is found
  Case else 
    response.write("ERROR")
End Select
'+++ +++ +++ +++ +++ +++ +++ +++ +++ +++ +++ +++ END CODE THAT SHOULD BE EDITED

'______________________________________________________________________________________ Write File
Do While objDataFilein.AtEndofStream <> True
'read the line at hand, store in mem
strRow = objDataFilein.ReadLine

if i =0 then    
    'if this is the header, add the column 'Hospital
    strRow = strHeader
Else     
    strRow = strRow & strSeparator & strHospitalCode
End if

'write the line
objFileout.WriteLine (strRow)  
i=i+1
Loop

'______________________________________________________________________________________ Close File
objDataFilein.Close
objFileout.Close

