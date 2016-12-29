VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPicScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' \\ -- Simple módulo para crear un Picturebox con barras de scroll
'    -- Autor : Luciano Lodola -- http://www.recursosvisualbasic.com.ar
' --------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------


' \\ -- Api para obtener medidas de los ScrollBar
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYHSCROLL                  As Long = 3
Private Const SM_CXVSCROLL                  As Long = 2

' \\ -- Variables de objeto para los controles
Private mOwnerForm                          As Form             ' -- Referencia al form
Private WithEvents pic_Imagen               As PictureBox       ' -- Referencia al control que contendrá la imagen
Attribute pic_Imagen.VB_VarHelpID = -1
Private WithEvents Pic_Container            As PictureBox       ' -- Referencia al contenedor
Attribute Pic_Container.VB_VarHelpID = -1
Private WithEvents Vertical_Scroll          As VScrollBar       ' -- Referencia al ScrollBar vertical
Attribute Vertical_Scroll.VB_VarHelpID = -1
Private WithEvents Horizontal_Scroll        As HScrollBar
Attribute Horizontal_Scroll.VB_VarHelpID = -1

' \\ -- Eventos
Event Resize()
Event VerticalScroll(ByVal iValue As Integer)
Event HorizontalScroll(ByVal iValue As Integer)
Event VerticalScrollChange()
Event HorizontalScrollChange()

' -------------------------------------------------------------------------------------------
' \\ -- Sub para cargar la imagen
' -------------------------------------------------------------------------------------------
Sub SetPicture(sPathImage As String)
    On Error GoTo error_Handler
    pic_Imagen.Picture = LoadPicture(sPathImage)
    Call Refresh
Exit Sub
error_Handler:
MsgBox Err.Description, vbCritical, "Error en SetPicture"
Call Refresh
End Sub
' -------------------------------------------------------------------------------------------
' \\ -- Sub para de inicio -- para referenciar los controles y configurarlos
' -------------------------------------------------------------------------------------------

Sub Init(PicBox_Container As PictureBox, OwnerForm As Form)
    
    ' -- referenciar
    Set Pic_Container = PicBox_Container
    Set mOwnerForm = OwnerForm
    Set Vertical_Scroll = Vertical_Scroll
    Set Horizontal_Scroll = Horizontal_Scroll
    
    ' -- Crear los controles en tiempo de ejecución
    With mOwnerForm.Controls
        Set pic_Imagen = .Add("vb.PictureBox", "pic_Imagen")
        Set Vertical_Scroll = .Add("vb.VScrollBar", "Vertical_Scroll")
        Set Horizontal_Scroll = .Add("vb.HScrollBar", "Horizontal_Scroll")
    End With
    
    ' -- Configurar pic_Imagen
    With pic_Imagen
        .AutoSize = True
        .BorderStyle = 0
        Set .Container = Pic_Container
    End With
    ' -- Configurar las barras
    With Horizontal_Scroll
        .LargeChange = 15
        Set .Container = Pic_Container      ' -- Insertar el objeto dentro del contenedor
        .ZOrder 0
    End With
    With Vertical_Scroll
        .LargeChange = 15
        Set .Container = Pic_Container
        .ZOrder 0                           ' -- TRaer el control al frente
    End With
    ' -- Refrescar
    Call Refresh
End Sub
' ----------------------------------------------------------------------------------------------
' \\ -- Refrescar ( Posicionar, ocultar, establecer la magnitud de cambio de los scroll etc ..)
' ----------------------------------------------------------------------------------------------
Sub Refresh()

    Dim sngVSCWidth         As Single
    Dim sngHSCHeight        As Single
    
    'Colocamos le pic_Imagen en la parte superior izquierda dentro del Pic_Container
    pic_Imagen.Move 0, 0
    
    ' -- Calcular ancho y alto de las barras
    With Screen
        sngVSCWidth = GetSystemMetrics(SM_CXVSCROLL) * .TwipsPerPixelX
        sngHSCHeight = GetSystemMetrics(SM_CYHSCROLL) * .TwipsPerPixelY
    End With
    
    
    ' -- Configurar control de imagen, máximos de las barras y posicionar
    With pic_Imagen
        Horizontal_Scroll.Max = (.Width - Pic_Container.ScaleWidth + Vertical_Scroll.Width) / 120 + 1
        Vertical_Scroll.Max = (.Height - Pic_Container.ScaleHeight + Horizontal_Scroll.Height) / 120 + 1
        If .Picture Is Nothing Or .Picture = 0 Then
            Horizontal_Scroll.Visible = False
            Vertical_Scroll.Visible = False
            .Visible = False
            .Move 0, 0, 0, 0
        Else
            .Visible = True
        End If
        Horizontal_Scroll.Visible = Not (.Width < Pic_Container.Width)
        Vertical_Scroll.Visible = Not (.Height < Pic_Container.Height)
    End With
    
    With Pic_Container
        Vertical_Scroll.Move .ScaleWidth - sngVSCWidth, 0, sngVSCWidth, .ScaleHeight
        If Vertical_Scroll.Visible Then
            Horizontal_Scroll.Move 0, .ScaleHeight - sngHSCHeight, .ScaleWidth - sngVSCWidth, sngHSCHeight
        Else
            Horizontal_Scroll.Move 0, .ScaleHeight - sngHSCHeight, .ScaleWidth, sngHSCHeight
        End If
    End With
End Sub
' ----------------------------------------------------------------------------------------------
' \\ -- Fin ( Destruir controles )
' ----------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    On Error GoTo error_Handler
    ' -- Eliminar los controles
    With mOwnerForm
        Call .Controls.Remove("pic_Imagen")
        Call .Controls.Remove("Vertical_Scroll")
        Call .Controls.Remove("Horizontal_Scroll")
    End With
    Exit Sub
error_Handler:
End Sub
' ----------------------------------------------------------------------------------------------
' \\ -- Eventos prrivados de controles ( Scroll, Change y Resize ) y lanzarlos con RaiseEvent
' ----------------------------------------------------------------------------------------------
Private Sub Horizontal_Scroll_Change()
    pic_Imagen.Left = -Horizontal_Scroll.Value * 120
    RaiseEvent HorizontalScrollChange
End Sub
Private Sub Horizontal_Scroll_Scroll()
    Horizontal_Scroll_Change
    RaiseEvent HorizontalScroll(Horizontal_Scroll.Value)
End Sub
Private Sub Pic_Container_Resize()
    Call Refresh
    RaiseEvent Resize
End Sub
Private Sub Vertical_Scroll_Change()
    pic_Imagen.top = -CSng(Vertical_Scroll.Value) * 120
    RaiseEvent VerticalScrollChange
End Sub
Private Sub Vertical_Scroll_Scroll()
    Vertical_Scroll_Change
    RaiseEvent VerticalScroll(Vertical_Scroll.Value)
End Sub
' -----------------------------------------------------------------------------------------------
' \\ -- Propiedades -- Value, max
' -----------------------------------------------------------------------------------------------
Property Get HSbarMax() As Long
    HSbarMax = Horizontal_Scroll.Max
End Property
Property Get VSbarMax() As Long
    VSbarMax = Vertical_Scroll.Max
End Property
Property Get VSBarValue() As Integer
    VSBarValue = Vertical_Scroll.Value
End Property
Property Let VSBarValue(iValue As Integer)
    If iValue <= Vertical_Scroll.Max And iValue > Vertical_Scroll.Min Then
        Vertical_Scroll.Value = iValue
    End If
End Property
Property Get HSBarValue() As Integer
    HSBarValue = Horizontal_Scroll.Value
End Property
Property Let HSBarValue(iValue As Integer)
    If iValue <= Horizontal_Scroll.Max And iValue > Horizontal_Scroll.Min Then
        Horizontal_Scroll.Value = iValue
    End If
End Property


