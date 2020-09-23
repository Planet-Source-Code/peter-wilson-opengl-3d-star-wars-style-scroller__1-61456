Attribute VB_Name = "mMain"
Option Explicit

Private m_frmCanvas As frmCanvas

Dim g_lngWindow As Long

Public Sub Main()

' Requirements.
' =============
'   * glut32.dll
'
' Also see:
' =========
'   * http://www.opengl.org/resources/faq/getting_started.html

    Set m_frmCanvas = New frmCanvas
    m_frmCanvas.Show
    
    Call DoInitOpenGL(m_frmCanvas.hDC)
    
    ' The glClearColor function specifies clear values for the color buffers.
    Call VBOpenGL.glClearColor(0, 0, 0, 0)
    
    ' The glClear function clears the buffer(s) to preset values.
    Call VBOpenGL.glClear(clrColorBufferBit Or clrDepthBufferBit)
    
    
    Call VBOpenGL.glViewport(0, 0, m_frmCanvas.ScaleWidth, m_frmCanvas.ScaleHeight)
    
    
    ' The glMatrixMode function specifies which matrix is the current matrix.
    Call VBOpenGL.glMatrixMode(mmProjection)
    
    

End Sub

