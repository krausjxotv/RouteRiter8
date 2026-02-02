Attribute VB_Name = "MWGFX"

'basic conversion and processing functions
Public Declare Function anytobmps _
               Lib "mwgfxvb.dll" _
               Alias "anytobmpsVB" (ByVal s As String, _
                                    ByVal d As String, _
                                    p As Pic, _
                                    ByVal A As Long, _
                                    ByVal B As Long) As Long
Public Declare Function anytogrey _
               Lib "mwgfxvb.dll" _
               Alias "anytogreyVB" (ByVal s As String, _
                                    ByVal d As String, _
                                    p As Pic, _
                                    ByVal A As Long, _
                                    ByVal B As Long) As Long
Public Declare Function anyto256 _
               Lib "mwgfxvb.dll" _
               Alias "anyto256VB" (ByVal s As String, _
                                   ByVal d As String, _
                                   p As Pic, _
                                   ByVal A As Long, _
                                   ByVal B As Long) As Long
Public Declare Function Pic_Convert _
               Lib "mwgfxvb.dll" _
               Alias "Pic_ConvertVB" (ByVal s As String, _
                                      ByVal d As String, _
                                      p As Pic, _
                                      ByVal A As Long, _
                                      ByVal B As Long, _
                                      ByVal c As Long, _
                                      ByVal d As Long) As Long
Public Declare Function bmptoanys _
               Lib "mwgfxvb.dll" _
               Alias "bmptoanysVB" (ByVal s As String, _
                                    ByVal d As String, _
                                    p As Pic, _
                                    ByVal A As Long, _
                                    ByVal B As Long) As Long
Public Declare Function checkbmp _
               Lib "mwgfxvb.dll" _
               Alias "checkbmpVB" (ByVal s As String, _
                                   p As Pic) As Long
Public Declare Function bmprocess _
               Lib "mwgfxvb.dll" _
               Alias "bmprocessVB" (ByVal s As String, _
                                    ByVal d As String, _
                                    p As Pic, _
                                    ByVal A As Long) As Long
'mwgfx24.dll functions (accessed through mwgfxvb.dll)
Public Declare Function WinImagePrint _
               Lib "mwgfxvb.dll" _
               Alias "WinImagePrintVB" (ByVal s As String) As Long
Public Declare Function WinImageBrowse _
               Lib "mwgfxvb.dll" _
               Alias "WinImageBrowseVB" (ByVal s As String) As Long
Public Declare Function WinImageCopy _
               Lib "mwgfxvb.dll" _
               Alias "WinImageCopyVB" (ByVal s As String) As Long
Public Declare Function WinImageSize _
               Lib "mwgfxvb.dll" _
               Alias "WinImageSizeVB" (ByVal s As String) As Long
Public Declare Function WinImageCrop _
               Lib "mwgfxvb.dll" _
               Alias "WinImageCropVB" (ByVal s As String, _
                                       ByVal x As Long, _
                                       ByVal Y As Long, _
                                       ByVal w As Long, _
                                       ByVal h As Long) As Long
Public Declare Function WinImageAdjust _
               Lib "mwgfxvb.dll" _
               Alias "WinImageAdjustVB" (ByVal s As String) As Long
Public Declare Function WinImageShow _
               Lib "mwgfxvb.dll" _
               Alias "WinImageShowVB" (ByVal s As String, _
                                       ByVal c As Long) As Long
Public Declare Function WinSlideShow _
               Lib "mwgfxvb.dll" _
               Alias "WinSlideShowVB" (ByVal s As String, _
                                       ByVal secs As Long, _
                                       ByVal locol As Long, _
                                       ByVal sloop As Long, _
                                       ByVal validext As Long, _
                                       ByVal sbeep As Long) As Long

