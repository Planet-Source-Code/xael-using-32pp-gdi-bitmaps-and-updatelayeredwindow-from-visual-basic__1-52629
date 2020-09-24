VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   0  'None
   Caption         =   "MainForm"
   ClientHeight    =   2400
   ClientLeft      =   11175
   ClientTop       =   1035
   ClientWidth     =   5280
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timFade 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4320
      Top             =   120
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' UpdateLayeredWindow Demonstration project
' Written by Yehovah, Yehowshua, and Isaac Umoh for Planet Source Code

' Feel free to use any and all portions of this code. Spread the joy :).

' NOTES:
' One thing you may want to remember is that in actual bitmaps (at least in DIB's), the color
' bits are actually arrayed like BGR(A), and not (A)RGB. Also, a device-independent bitmap
' has an origin in the lower-left corner of the screen (for the longest time I thought it
' has the lower-right, Thank God for correcting me in this), while a device-dependent bitmap
' has an origin at the upper-left.

' Also, although I did not implement it here, in a production-quality app, you should
' subclass the window and handle the WM_DISPLAYCHANGE event. When you receive this
' event you should recreate any surfaces that you plan to send to UpdateLayeredWindow

' API Declarations

' Consts
Private Const DEF_APP_TRANSPARENCY = 192   ' You can play around with this value
                                            ' to make the background image more
                                            ' or less transparent.
Private Const vbLongSize As Long = 2147483647
Private Const ULW_OPAQUE = &H4      ' Used to tell UpdateLayeredWindow to draw the window without alpha-blending
Private Const ULW_COLORKEY = &H1    ' Used to draw the window with a color key
Private Const ULW_ALPHA = &H2       ' Used to draw the window with alpha-blending.
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

' Types

' Used by RndFunction
Private Type UUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

' Used by UpdateLayeredWindow and AlphaBlend
Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

' Used by ARGB, and GetARGB
Private Type Color
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

' Used to set the positions for our balls of light
Private Type LightBall
    LightX As Integer
    LightY As Integer
    XVel As Integer
    YVel As Integer
    Alpha As Integer
    AlphaIncrement As Integer
End Type


' Not really used
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

' Function declarations
Private Declare Function UuidCreate Lib "rpcrt4.dll" (ByRef rUUID As UUID) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Variable Declarations
Private m_mainSurfaceDC As Long
Private m_mainSurfaceBitmap As Long

Private m_backSurfaceDC As Long
'Private m_bufferSurfaceBitmap As Long

Private m_lightMapDC As Long
Private m_lightMapBitmap As Long
Private m_lmWidth As Long
Private m_lmHeight As Long

Private m_backPic As StdPicture
Dim m_backHeight As Long
Dim m_backWidth As Long

Private m_blendFunc As BLENDFUNCTION
Private m_backBlendFUnc As Long

Private m_windowSize As Size
Private m_srcPoint As POINTAPI


Private m_lBlendFunc As Long

Private m_lightRows As Integer
Private m_lightColumns As Integer

Private m_lightBalls(0 To 4) As LightBall

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Initialize()
    Dim tempBI As BITMAPINFO            ' Holds the bitmap information
    Dim tempBlend As BLENDFUNCTION      ' Used to specify what kind of blend we want to perform
    Dim lBlendFunc As Long
    'Dim tempBits() As Long              ' Holds an array of bits from a bitmap surface
    
    Dim I As Long
    
    ' Let's tell windows that we want to make this a layered window
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    
    ' Make the window a top-most window so we can always see the cool stuff
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    ' One primary difference between UpdateLayeredWindow and SetLayeredWindowAttributes
    ' is that when using the former, the Paint event is not used. This is because you
    ' use a bitmap buffer to perform any drawing that your need to do, and then
    ' pass this buffer to window using UpdateLayeredWindow. Windows then handles
    ' the actual display of the window and all hit-testing.
    
    ' Ok, here's what this we are going to do
    ' 1. We need to read the light map bitmap (LightMap3.bmp) which is 24 bits, and
    '    convert it to a 32 bits, by changing the fading to black, to white pixels with
    '    varying alpha values. This light map, will then be used to create some balls
    '    of light which will move around the program surface, and
    '    demonstrate per-pixel alpha blending.
    
    ' 2. Next we'll load up the background surface that will form the application's
    '    form. It's a 24bit bitmap.
    
    ' 3. Next we need to create a 32bpp surface to use as the main application surface which
    '    we will pass toUpdateLayeredWindow. This needs to be a 32bpp surface because
    '    it's going to be semi-transparent, so we want the light values from the light-map
    '    to blend in with the windows desktop.
    
    ' 4. Finally we'll draw the background surface onto the main application surface by
    '    using the Alpha Blend function to make it semi-transparent; and also initialize
    '    an array of "LightBall" objects for each of our semi-transparent light balls.
    
    ' 5. After calling UpdateLayeredWindow with 32bpp surface, we will start a timer
    '    that will move each of the light balls around the application surface. This
    '    is accomplished by basically clearing the main application surface, re-drawing
    '    the 24bit background semi-transparently, and finally drawing each of the light balls.
    '    When we are done, we call UpdateLayeredWindow so that it can update the onscreen
    '    image.
    
    ' First let's create the light map surface
    BuildLightMap
    
    ' Let's create two DC's; one for our 32bpp surface, and another for the
    ' background surface, which we will load from file.
    m_mainSurfaceDC = CreateCompatibleDC(Me.hdc)
    m_backSurfaceDC = CreateCompatibleDC(Me.hdc)
    
    ' Now let's load up the background surface (24 bit)
    Set m_backPic = LoadPicture(App.Path & "\Title.bmp")
    m_backHeight = ScaleY(m_backPic.Height, vbHimetric, vbPixels)
    m_backWidth = ScaleX(m_backPic.Width, vbHimetric, vbPixels)
    SelectObject m_backSurfaceDC, m_backPic.handle
    
    ' Now let's create the DIB Section. A DIB Section is basically just device independent
    ' bitmap that we can write to. We've mainly using this function because we
    ' need to create a 32 bit bitmap as an main surface
    
    ' First we need to setup the BitmapInfo struct which tells windows
    ' about the structure of the bitmap we are creating
    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32    ' Each pixel is 32 bit's wide
        .biHeight = Me.ScaleHeight  ' Height of the form
        .biWidth = Me.ScaleWidth    ' Width of the form
        .biPlanes = 1   ' Always set to 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
    End With
    m_mainSurfaceBitmap = CreateDIBSection(m_mainSurfaceDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    SelectObject m_mainSurfaceDC, m_mainSurfaceBitmap   ' Select the new bitmap
                                                        ' into our DC.
                                                        
    ' Now the above bitmap by default has it's bit's initialized to 0, which means
    ' that while the color of all of the pixels is black, the alpha bit for the pixels
    ' is 0, therefore the entire image is transparent. With a 32bpp surface, color functions
    ' such as RGB will not work. Also, you cannot blit pictures directly to the surface
    ' using functions like BitBlt, has the alpha bit WILL NOT be set, and the image is
    ' technically transparent. Also, when using per-pixel alpha values, all of the color
    ' bits of the image must be "premultipled" with the alpha value and then divided by
    ' 255. For example, let's say you have a bitmap where all of the pixels have an Alpha
    ' value of 128, and the Red bit field is set to 255. In order for this surface to be
    ' displayed properly, you need to loop through all of the bits in the image and set
    ' the red bit to (255/128)*255.
    
    ' The easiest way to blit a 24bit surface onto a 32bit surface is to use the AlphaBlend
    ' function with the SourceConstantAlpha value set to any value besides 255. Then
    ' AlphaBlend will automatically convert the 24bit surface to a 32bit surface and
    ' premultiply the pixels for you. If you set the SourceConstantAlpha to 255, then
    ' AlphaBlend will basically just behave like a normal BitBlt, and you'll get a mess.
    ' Go ahead, and try it, to see what I mean.
    
    ' First create a BlendFunction which tells AlphaBlend how to blend the source
    ' and destination pixels
    With tempBlend
        .AlphaFormat = 0    ' This field specifies whether or not the source image is a
                            ' 32bit bitmap. Here we set it to 0 to indicate that it's not.
        .BlendFlags = 0     ' Always set to 0
        .BlendOp = AC_SRC_OVER  ' Specifies the type of blending operation to perform.
                                ' Currently this is the only supported one.
        .SourceConstantAlpha = DEF_APP_TRANSPARENCY  ' Specifies the "opacity" of the source image. 255 means
                                    ' that the entire image is opaque, while 0 means that
                                    ' the entire image is transparent. All of the values
                                    ' in-between specify different levels of transparency.
                                    ' Here we want the background image, to be partially
                                    ' transparent, so we use 192.
    End With
    
    ' Save a global copy (which we will use later in the Timer (trying to save some cycles :) )
    CopyMemory m_backBlendFUnc, tempBlend, 4
    
    ' Now let's blend the background surface onto the main surface
    AlphaBlend m_mainSurfaceDC, 0, (Me.ScaleHeight - m_backHeight) / 2, Me.ScaleWidth, m_backHeight, m_backSurfaceDC, 0, 0, Me.ScaleWidth, m_backHeight, m_backBlendFUnc
    
    ' Now it's time to initialize our light ball array which holds the position
    ' and velocity of each ball.
    
    ' Now let's determine the maximum bounds for our light balls
    m_lightRows = (m_backWidth / 32) - 1
    m_lightColumns = (m_backHeight / 32) - 1
    
    For I = 0 To UBound(m_lightBalls)
        With m_lightBalls(I)
            ' We're setting the values to this, so that the code
            ' in the timer will randomly select positions and speeds for each of the
            ' balls.
            .LightX = Me.ScaleWidth '-m_lmWidth
            .LightY = Me.ScaleHeight '(-(m_lmHeight / 2) + (32 * 2)) + 1
            .XVel = 0
            .YVel = 0
            '.Alpha = 255
            '.AlphaIncrement = 1
        End With
    Next I
    
    ' Now it's time to call UpdateLayeredWindow to display this on the desktop
    ' The basic syntax consists of passing the window handle, the window device context,
    ' the source coordinates, and the size of the window ALONG with the DC which contains
    ' all of the information we want shown onscreen. In order to implement transparency
    ' you can use a color key, to turn all of the pixels of that color transparent, or
    ' you can use a BLENDFUNCTION in a way very similar to the AlphaBlend function
    ' call above. If you implement code to move the window, you shound call
    ' UpdateLayeredWindow each time the position changes, and pass a POINTAPI struct holding
    ' the new coordinates (pptDst).
    
    ' First let's setup the window size, and the source coordinates
    ' that we want to use in the layer. We are using global variables
    ' so that we can use this pre-initialized values in the timer loop.
    m_srcPoint.x = 0
    m_srcPoint.y = 0
    m_windowSize.cx = Me.ScaleWidth
    m_windowSize.cy = Me.ScaleHeight
    
    ' Use Alpha (ULW_ALPHA)
    With m_blendFunc
        .AlphaFormat = AC_SRC_ALPHA  ' Now we sent this to AC_SRC_ALPHA since our bitmap
                                    ' is 32 bits.
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    CopyMemory m_lBlendFunc, m_blendFunc, 4
    UpdateLayeredWindow Me.hwnd, Me.hdc, ByVal 0, m_windowSize, m_mainSurfaceDC, m_srcPoint, 0, m_blendFunc, ULW_ALPHA
    timFade.Enabled = True
    
    ' Use Color Key (ULW_COLORKEY)
    'UpdateLayeredWindow Me.hwnd, Me.hdc, ByVal 0, windowsize, m_mainSurfaceDC, srcpoint, RGB(255, 0, 255), tempBlend, ULW_COLORKEY
    
    ' Normal Blt (ULW_OPAQUE)
    'UpdateLayeredWindow Me.hwnd, Me.hdc, ByVal 0, windowsize, m_mainSurfaceDC, srcpoint, 0, tempBlend, ULW_OPAQUE
    
    ' There that's it. Now onscreen we have a semi-transparent form (yes folks, you can really
    ' see the desktop back there). Now onto the timer loop to add those cool light balls.
    ' :)
    
    ' BTW, as an additional programming excerise why don't you
    ' try to add a gradient to the background so that it looks
    ' like it's slowly fading to the desktop. Haha, or you
    ' could try to add an alpha shadow, maybe add support for an
    ' anti-aliased circle surface (shouldn't be to hard if you can
    ' find some code to load a 32bit PNG file.
End Sub

Private Sub Form_Load()
    ' Note: DO NOT try to call UpdateLayeredWindows from here. It will return 0 (:?)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ' Destroy the objects
    DeleteObject m_mainSurfaceBitmap
    DeleteDC m_mainSurfaceDC
    
    DeleteObject m_lightMapBitmap
    DeleteDC m_lightMapDC
    
    Set m_backPic = Nothing
    DeleteDC m_backSurfaceDC
End Sub

' Feel free to optimize this up anyway you see fit.
Private Sub timFade_Timer()
    Dim rndValue As Long
    Dim lBlend As Long
    Dim I As Integer
    
    ' First clean the surface (kinda costly)
    ClearSurface m_mainSurfaceDC, m_mainSurfaceBitmap, Me.ScaleWidth, Me.ScaleHeight
    
    ' Now ReDraw the background image onto the surface
    AlphaBlend m_mainSurfaceDC, 0, (Me.ScaleHeight - m_backHeight) / 2, Me.ScaleWidth, m_backHeight, m_backSurfaceDC, 0, 0, Me.ScaleWidth, m_backHeight, m_backBlendFUnc
    
    ' Now it's time to draw the light balls (yippee! :)))
    For I = 0 To UBound(m_lightBalls)
        With m_lightBalls(I)
            .LightX = .LightX + .XVel
            .LightY = .LightY + .YVel
            
            If .LightX >= Me.ScaleWidth Or .LightY >= Me.ScaleHeight Then
                ' Determine how to draw the the pixel this time (either row based or column based
                rndValue = RndFunction(1, 0)
                If rndValue = 0 Then
                    ' Draw row based
                    .LightX = -m_lmWidth
                    .LightY = (-(m_lmHeight / 2) + (32 * RndFunction(m_lightColumns, 2))) - 19
                    .XVel = RndFunction(8, 1)
                    .YVel = 0
                    
                    '.AlphaIncrement = -.XVel
                    '.Alpha = 255
                Else
                    ' Draw Column based
                    .LightX = (-(m_lmWidth / 2) + (32 * RndFunction(m_lightRows, 2))) - 19 '- (Me.ScaleHeight - m_backHeight)
                    .LightY = -m_lmHeight
                    .XVel = 0
                    .YVel = RndFunction(8, 1)
                    
                    '.AlphaIncrement = -.YVel
                    '.Alpha = 255
                End If
            End If
            
            ' ---  NOT USED (You can trun it on if you'd like, it just oscillates the alpha
            ' channels of each of the balls [Be sure to uncomment the other lines as well])
            '.Alpha = .Alpha + .AlphaIncrement
            'If .Alpha >= 255 Then
            '    .Alpha = 255
            '    .AlphaIncrement = -.AlphaIncrement
            'ElseIf .Alpha <= 0 Then
            '    .Alpha = 0
            '    .AlphaIncrement = -.AlphaIncrement
            'End If
            ' Draw the light ball (note: you could use an array lookup to speed this up
            'm_blendFunc.SourceConstantAlpha = .Alpha
            'CopyMemory lBlend, m_blendFunc, 4
            
            ' Go ahead and draw the ball onto the surface
            AlphaBlend m_mainSurfaceDC, .LightX, .LightY, m_lmWidth, m_lmHeight, m_lightMapDC, 0, 0, m_lmWidth, m_lmHeight, m_lBlendFunc
        End With
    Next I
    
    ' Now update the window (:>)
    'm_blendFunc.SourceConstantAlpha = 255
    UpdateLayeredWindow Me.hwnd, Me.hdc, ByVal 0, m_windowSize, m_mainSurfaceDC, m_srcPoint, 0, m_blendFunc, ULW_ALPHA
End Sub

' ---- SUPPORT FUNCTIONS ----
' The following functions are used to make the entire project. My thanks to God my father,
' his Son Yehowshua Christ for helping me develop each of them.

' An improved randomization function which uses UUID to generate unrelated sequences of
' random numbers.
Public Function RndFunction(ByVal HighValue As Long, ByVal LowValue As Long) As Integer
    Dim tempUUID As UUID
    
    UuidCreate tempUUID
    
    RndFunction = Int((HighValue - LowValue + 1) * Abs(tempUUID.Data1 / vbLongSize) + LowValue)
    'RndFunction = Int((HighValue - LowValue + 1) * Rnd + LowValue)
End Function

' Used to build an ARGB Long (note: If you read my last article on making a custom RGB
' function you may be wondering why this here uses CopyMemory instead of the bit masks.
' I was not able to get alpha channel to work uses the other method, it kept causing an
' overflow. I'll work on it later, and post up my results, God willing.

' By the way, this function can automatically premulitply the colors, so that you can
' apply in directly to a surface without doing it yourself (you will, however have to
' create a custom function to do the filling as the standard GDI fill functions DO NOT
' work well with 32bpp surfaces. GetARGB automatically demultiplys the values so that they
' will read correctly.
Private Function ARGB(ByVal Alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, Optional ByVal preMultiply As Boolean = True) As Long
    Dim tempColor As Color
    Dim dPre As Double
    Dim lColor As Long
    
    If preMultiply Then
        dPre = Alpha / 255
    Else
        dPre = 1
    End If
    With tempColor
        .Alpha = Alpha '192
        .Red = Red * dPre '78 'CLng(128 * tempColor.Alpha) / 255
        .Green = Green * dPre
        .Blue = Blue * dPre
    End With
    dPre = 0
    CopyMemory lColor, tempColor, 4
    ARGB = lColor
End Function

' Used to get the individual bits from an ARGB value
Private Function GetARGB(ByVal lColor As Long, ByRef Alpha As Byte, ByRef Red As Byte, ByRef Green As Byte, ByRef Blue As Byte, Optional demultiply As Boolean = True)
    Dim tempColor As Color
    
    CopyMemory tempColor, lColor, 4
    
    Alpha = tempColor.Alpha
    If demultiply And tempColor.Alpha <> 0 Then
        Red = Round(CDbl(tempColor.Red / Alpha) * 255, 0)
        Green = Round(CDbl(tempColor.Green / Alpha) * 255, 0)
        Blue = Round(CDbl(tempColor.Blue / Alpha) * 255, 0)
    End If
End Function


' This function is used to clear a 32bpp surface (make it transparent)
Private Sub ClearSurface(ByVal DC As Long, ByVal Bitmap As Long, ByVal Width As Long, ByVal Height)
    Dim tempBits() As Long
    Dim tempBI As BITMAPINFO
    Dim I As Long
    
    ' Setup the into struct
    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32    ' Each pixel is 32 bit's wide
        .biHeight = Height  ' Height
        .biWidth = Width    ' Width
        .biPlanes = 1   ' Always set to 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
    End With
    ReDim tempBits(Width * Height)
    
    ' Grab the bits
    GetDIBits DC, Bitmap, 0, Height, tempBits(0), tempBI, DIB_RGB_COLORS
    For I = 0 To UBound(tempBits)
        ' Sets all of the bits to Alpha=0, Red=0, Green=0, Blue=0
        tempBits(I) = 0
    Next I
    
    ' Put em' back
    SetDIBits DC, Bitmap, 0, Height, tempBits(0), tempBI, DIB_RGB_COLORS
    Erase tempBits
End Sub

' This is used to generate a 32bpp bitmap light-map from a 24-bit one.
' Believe or not, code similar to this is what You'll have to do to draw
' good looking text on an 32ppp surface. (Unless of course you'd like to add your
' OWN anti-aliasing code Yuck :( )
Private Sub BuildLightMap()
    Dim tempPic As StdPicture
    Dim tempDC As Long
    
    Dim tempBI As BITMAPINFO
    Dim tempBlend As BLENDFUNCTION
    Dim lBlend As Long
    
    Dim tempBits() As Long
    Dim I As Long
    Dim tempAlpha As Byte, tempRed As Byte, tempGreen As Byte, tempBlue As Byte
    Dim tempDbl As Double
    
    ' Load up the light map (made in PhotoImpact)
    tempDC = CreateCompatibleDC(Me.hdc)
    Set tempPic = LoadPicture(App.Path & "\LightMap3.bmp")
    SelectObject tempDC, tempPic
    
    ' Get the dimensions
    m_lmWidth = ScaleX(tempPic.Width, vbHimetric, vbPixels)
    m_lmHeight = ScaleY(tempPic.Height, vbHimetric, vbPixels)
        
     ' Create the surface
    m_lightMapDC = CreateCompatibleDC(Me.hdc)
    
    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32    ' Each pixel is 32 bit's wide
        .biHeight = m_lmHeight
        .biWidth = m_lmWidth
        .biPlanes = 1   ' Always set to 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
    End With
    ReDim tempBits(m_lmWidth * m_lmHeight)
    m_lightMapBitmap = CreateDIBSection(m_lightMapDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    SelectObject m_lightMapDC, m_lightMapBitmap
    
    ' Blend it with the surface (AlphaBlend 254)
    With tempBlend
        .AlphaFormat = 0
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 254
    End With
    CopyMemory lBlend, tempBlend, 4
    AlphaBlend m_lightMapDC, 0, 0, m_lmWidth, m_lmHeight, tempDC, 0, 0, m_lmWidth, m_lmHeight, lBlend
    
    Set tempPic = Nothing
    DeleteDC tempDC
    
    ' Now convert all of the semi-white pixels to white pixels that are semi-transparent
    GetDIBits m_lightMapDC, m_lightMapBitmap, 0, m_lmHeight, tempBits(0), tempBI, DIB_RGB_COLORS
        For I = 0 To UBound(tempBits)
            ' Get the surface color into
            GetARGB tempBits(I), tempAlpha, tempRed, tempGreen, tempBlue
        
            ' Destroy any black pixels
            If tempRed = 0 And tempGreen = 0 And tempBlue = 0 Then
                tempBits(I) = 0
            ElseIf Not (tempRed = 254 And tempGreen = 254 And tempBlue = 254) And Not (tempRed = 255 And tempGreen = 255 And tempBlue = 255) And tempAlpha = 254 Then
                ' Calculate the the percentage from white by calculating the percentages
                ' for each of the colors and taking the average. (also seems to work THANK GOD!)
                tempDbl = CDbl(CDbl(tempRed / 254) + CDbl(tempGreen / 254) + CDbl(tempBlue / 254)) / 3
                tempBits(I) = ARGB((255 * Round(tempDbl, 2)), 254, 254, 254)
            End If
        Next I
        
    ' Set the bits back
    SetDIBits m_lightMapDC, m_lightMapBitmap, 0, m_lmHeight, tempBits(0), tempBI, DIB_RGB_COLORS
    Erase tempBits
    
    ' All Done
End Sub

