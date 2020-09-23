VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form frmCanvas 
   BorderStyle     =   0  'None
   Caption         =   "OpenGL v1.0"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   8475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MCI.MMControl MMControl1 
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   979
      _Version        =   393216
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Copyright © 2003-2005 Peter Wilson (peter@sourcecode.net.au)
' ============================================================

Private m_hglrc As Long

' Display List Base
Private m_DisplayListBase As VBOpenGL.GLuint

Private Sub Init_MIDISoundtrack()

    ' ==========================================
    ' Init. Multimedia Control & Open MIDI File.
    ' ==========================================
    With Me.MMControl1
        .Visible = False
        .DeviceType = "Sequencer"
        .FileName = App.Path & "\The last Mohican.mid"
        .UpdateInterval = 1
        .Command = "OPEN"
        .Command = "PLAY"
    End With
    
End Sub
Public Sub DoInitOpenGL(hdc As Long)
        
    Dim pfd             As VBOpenGL.PIXELFORMATDESCRIPTOR
    Dim iPixelFormat    As Long
    
    ' ====================================================================================
    ' The PIXELFORMATDESCRIPTOR structure describes the pixel format of a drawing surface.
    ' ------------------------------------------------------------------------------------
    ' PFD_DRAW_TO_WINDOW        The buffer can draw to a window or device surface.
    ' PFD_SUPPORT_OPENGL        The buffer supports OpenGL drawing
    ' PFD_TYPE_RGBA             RGBA pixels. Each pixel has four components in
    '                           this order: red, green, blue, and alpha.
    ' ====================================================================================
    ZeroMemory pfd, Len(pfd)
    pfd.nSize = Len(pfd)
    With pfd
        .nVersion = 1
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
        .iPixelType = PFD_TYPE_RGBA
        .cDepthBits = 32
        .iLayerType = PFD_MAIN_PLANE
    End With
    
    
    ' ===========================================================================
    ' The ChoosePixelFormat function attempts to match an appropriate pixel
    ' format supported by a device context to a given pixel format specification.
    ' ===========================================================================
    iPixelFormat = VBOpenGL.ChoosePixelFormat(hdc, pfd)
    If iPixelFormat = 0 Then
        Call Err.Raise(Err.Number, "mOpenGL.DoInitOpenGL", "Could not choose the pixel format. (Try a different format)")
    End If
    
    
    ' =================================================================================
    ' The SetPixelFormat function sets the pixel format of the specified device context
    ' to the format specified by the iPixelFormat index.
    ' =================================================================================
    If VBOpenGL.SetPixelFormat(hdc, iPixelFormat, pfd) = False Then
        Call Err.Raise(Err.Number, "mOpenGL.DoInitOpenGL", "Could not set the pixel format.")
    End If
    
    
    ' ======================================================================================
    ' Create a rendering context.
    ' ---------------------------
    ' Note: Set the pixel format of the device context before creating a rendering context.
    ' ======================================================================================
    m_hglrc = VBOpenGL.wglCreateContext(hdc)
    If m_hglrc = 0 Then
        Call Err.Raise(Err.Number, "mOpenGL.DoInitOpenGL", "Could not create a rendering context.")
    End If


    ' ==========================================================================
    ' The wglMakeCurrent function makes a specified OpenGL rendering context the
    ' calling thread's current rendering context. All subsequent OpenGL calls
    ' made by the thread are drawn on the device identified by hdc.
    ' ==========================================================================
    If VBOpenGL.wglMakeCurrent(hdc, m_hglrc) = False Then
        Call Err.Raise(Err.Number, "mOpenGL.DoInitOpenGL", "Could not make the rendering context current.")
    End If


End Sub


Private Sub glPrintText(Text As String)
    
    Dim varWords    As Variant
    Dim intW        As Integer
    Dim intN        As Integer
    Dim intLen      As Integer
    Dim Bytes()     As Byte
    
    ' Clean up
    Text = Trim(Text)
    
    ' Split
    varWords = Split(Text, " ")
    
    
    Dim dblScale As Double
    dblScale = (1 / 4096)
    Call VBOpenGL.glPushMatrix
    
        Call VBOpenGL.glScaled(dblScale, dblScale, dblScale)
    
        ' Loop through all words
        For intW = LBound(varWords) To UBound(varWords)
        
            ' Get length of text.
            intLen = Len(varWords(intW))
                
            ' Size array
            ReDim Bytes(intLen - 1) As Byte
            
            Call VBOpenGL.glListBase(-33 + m_DisplayListBase)
            
            For intN = 0 To intLen - 1
                Bytes(intN) = CByte(Asc(Mid(varWords(intW), intN + 1, 1)))
            Next intN
            
            ' Draw Word
            Call VBOpenGL.glCallLists(intLen, VBOpenGL.GL_UNSIGNED_BYTE, Bytes(0))
        
            ' Draw Space between words (since there is no character for a space)
            Call VBOpenGL.glTranslatef((1 / dblScale) * 3, 0, 0)
        Next intW
        
    Call VBOpenGL.glPopMatrix
    
    ' Line Feed
    Call VBOpenGL.glTranslated(0, -7, 0)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    Call DoInitOpenGL(Me.hdc)
    
    m_DisplayListBase = DoCreateFonts(Me)
    
    Call Init_MIDISoundtrack
    
End Sub


Private Sub DoRefresh(Position As Long)
    
    ' Draws the Scence.
    
    Static dblD         As Double
    Static dblA         As Double
    Static dblB         As Double
    
    Dim dblAspectRatio  As Double
    Dim dblScale        As Double
    
    ' Let the user resize the window, and maintain the correct aspect ratio.
    dblAspectRatio = ((Me.ScaleWidth - 64) / (Me.ScaleHeight - 64))
    
    ' The glClearColor function specifies clear values for the color buffers.
    Call VBOpenGL.glClearColor(0, 0, 0, 0)

    ' Clear the buffers.
    Call VBOpenGL.glClear(clrColorBufferBit Or clrDepthBufferBit)

    ' The glMatrixMode function specifies which matrix is the current matrix.
    Call VBOpenGL.glMatrixMode(mmProjection)
    Call VBOpenGL.glLoadIdentity
    
    ' MatrixViewMapping_Per (This is a nice one to change).
    Call VBOpenGL.gluPerspective(90, dblAspectRatio, 0.1, 196)
    
    ' ViewPort
    Call VBOpenGL.glViewport(0, 0, Me.ScaleWidth, Me.ScaleHeight)
    
    ' Enable Z-Buffer.
    Call VBOpenGL.glEnable(glcDepthTest)
    
    ' Enable Fog.
    Call VBOpenGL.glFogf(fogMode, GL_LINEAR)
    Call VBOpenGL.glFogf(fogStart, 32)
    Call VBOpenGL.glFogf(fogEnd, 196)
    Call VBOpenGL.glFogf(fogDensity, 1)
    Dim objFogColour(3) As VBOpenGL.GLfloat
    objFogColour(0) = 0
    objFogColour(1) = 0
    objFogColour(2) = 0.2
    objFogColour(3) = 0
    Call VBOpenGL.glFogfv(fogColor, objFogColour(0))
    Call VBOpenGL.glEnable(glcFog)
    
    
    Call VBOpenGL.glMatrixMode(mmModelView)
    Call VBOpenGL.glLoadIdentity
    
    ' View Orientation (This is a nice one to change)
    dblD = dblD - 0.3
    Call VBOpenGL.gluLookAt(93, dblD, 93, 93, dblD + 64, 0, 0, 1, 0)
    
    
    Call VBOpenGL.glTranslated(0, -30, 0)
    Call VBOpenGL.glColor3f(0.3, 1, 1)
    
    dblScale = 3
    Call VBOpenGL.glPushMatrix
        Call VBOpenGL.glScaled(dblScale * 1.4, dblScale, dblScale)
        Call VBOpenGL.glLineWidth(2)
        Call glPrintText("Karmic Wars")
        Call VBOpenGL.glLineWidth(1)
    Call VBOpenGL.glPopMatrix
    
    Call VBOpenGL.glPushMatrix
        Call VBOpenGL.glTranslated(0, -10, 0)
        Call glPrintText("EPIDOSE IV - SUPREME JUSTICE")
        Call glPrintText("(prequels to follow)")
        Call VBOpenGL.glTranslated(0, -30, 0)
        Call glPrintText("They've always been here... They always will...")
        Call VBOpenGL.glTranslated(0, -30, 0)
        Call glPrintText("The sky setting, the stars alight -")
        Call glPrintText("and then there came a wonderous sight,")
        Call glPrintText("The stars - they moved and danced about,")
        Call glPrintText("with smiling waves, we welcomed them down.")
        Call VBOpenGL.glTranslated(0, -30, 0)
        Call glPrintText("A gasp went out amongst the crowd -")
        Call glPrintText("vapourized... then a mushroom cloud.")
        Call VBOpenGL.glTranslated(0, -30, 0)
        Call glPrintText("A priest set apart from the rest - ")
        Call glPrintText("sounded them out, to welcome the guest.")
        Call VBOpenGL.glTranslated(0, -15, 0)
        Call glPrintText("""Can there be peace?"", he asked.")
        Call VBOpenGL.glTranslated(0, -15, 0)
        Call glPrintText("For one brief moment tears seemed to well in the")
        Call glPrintText("beasts eyes.")
        Call VBOpenGL.glTranslated(0, -15, 0)
        Call glPrintText("He sighed out loud and his large chest fell -")
        Call glPrintText("His head - not proud - but rather sunken,")
        Call glPrintText("he spoke into his chest or maybe the ground...")
        Call VBOpenGL.glTranslated(0, -50, 0)
        Call glPrintText("""the wicked need a place to live too...""")
        Call VBOpenGL.glTranslated(0, -60, 0)
        
        Call VBOpenGL.glColor3f(1, 1, 0)
        Call glPrintText("Get ready to vote your brains out....")
        Call glPrintText("Not just a game, but a story and lesson for us all...")
        Call VBOpenGL.glTranslated(0, -15, 0)
        Call glPrintText("Coming soon to an open source code community near you!")
        Call VBOpenGL.glTranslated(0, -15, 0)
        Call glPrintText("Peter Wilson")
        Call glPrintText("peter@sourcecode.net.au")
        Call VBOpenGL.glTranslated(0, -190, 0)
        Call VBOpenGL.glColor3f(0, 0, 1)
        Call glPrintText("Secret Code: Blue Orb")
    Call VBOpenGL.glPopMatrix

    ' Update display.
    Call VBOpenGL.glFlush
    Call VBOpenGL.SwapBuffers(Me.hdc)
    
    If dblD < -685 Then
        Unload Me
    End If
    
End Sub

Private Sub MMControl1_StatusUpdate()

    ' Do Animation Loop
    Call DoRefresh(Me.MMControl1.Position)
    
End Sub


