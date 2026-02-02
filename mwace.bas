Attribute VB_Name = "mwace"
Type Pic
    width As Long
    height As Long
    depth As Long
    numcols As Long
    fin As Long
    comment(80) As Byte
    cmap(768) As Byte
    jlib As Long
    tlib As Long
    plib As Long
    buff As Long
    fout As Long
    ptr As Long
    progress As Long
    ilib As Long
    stype As Long
    spare(120) As Byte
End Type
    
Public Declare Function AceToBmp _
               Lib "mwacevb.dll" _
               Alias "AceToBmpVB" (ByVal s As String, _
                                   ByVal d As String) As Long
Public Declare Function AceToBmps _
               Lib "mwacevb.dll" _
               Alias "AceToBmpsVB" (ByVal s As String, _
                                    ByVal d As String, _
                                    ByVal A As String) As Long
Public Declare Function AceToTga _
               Lib "mwacevb.dll" _
               Alias "AceToTgaVB" (ByVal s As String, _
                                   ByVal d As String) As Long
Public Declare Function AceToTgaSquare _
               Lib "mwacevb.dll" _
               Alias "AceToTgaSquareVB" (ByVal s As String, _
                                         ByVal d As String) As Long
Public Declare Function BmpsToTga _
               Lib "mwacevb.dll" _
               Alias "BmpsToTgaVB" (ByVal s As String, _
                                    ByVal d As String, _
                                    ByVal A As String) As Long
Public Declare Function BmpsToTgaSquare _
               Lib "mwacevb.dll" _
               Alias "BmpsToTgaSquareVB" (ByVal s As String, _
                                          ByVal d As String, _
                                          ByVal A As String) As Long
Public Declare Function CheckAce _
               Lib "mwacevb.dll" _
               Alias "CheckAceVB" (ByVal s As String, _
                                   p As Pic) As Long

