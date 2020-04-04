Attribute VB_Name = "Module1"

Public Function Predict(SurfLat As Double, SurfLong As Double, LatLength As Double, Fluid As Double, Proppant As Double, Stages As Integer)
    Dim hReq As Object, JSON As Dictionary
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
    myurl = "http://54.184.11.112/api/MLModels/5c6c2712ec264e5720d26f4b/predict"
    xmlhttp.Open "PATCH", myurl, False
    xmlhttp.setRequestHeader "Content-Type", "text/json"
    xmlhttp.Send "[{'parameterName': 'Surface Latitude','value': '" & CStr(SurfLat) & "'},{'parameterName': 'Surface Longitude','value': '" & CStr(SurfLong) & "'},{'parameterName': 'Lat Length','value': '" & CStr(LatLength) & "'},{'parameterName': 'Total Fluid (gals)','value': '" & CStr(Fluid) & "'},{'parameterName': 'Total Proppant (lbs)','value': '" & CStr(Proppant) & "'},{'parameterName': 'NumStages','value': '" & CStr(Stages) & "'}]"
    Predict = Left(Mid(xmlhttp.responseText, 10, 100), InStr(Mid(xmlhttp.responseText, 10, 100), ",") - 1)
End Function

