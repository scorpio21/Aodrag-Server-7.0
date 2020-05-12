VERSION 5.00
Begin VB.Form FrmInterv 
   Caption         =   "Intervalos"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVampire 
      Height          =   285
      Left            =   1680
      TabIndex        =   57
      Text            =   "0"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox TxtFlechas 
      Height          =   285
      Left            =   120
      TabIndex        =   56
      Text            =   "0"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar Intervalos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   48
      Top             =   6480
      Width           =   3255
   End
   Begin VB.Frame Frame12 
      Caption         =   "Clima && Ambiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4875
      TabIndex        =   42
      Top             =   2130
      Width           =   2865
      Begin VB.Frame Frame7 
         Caption         =   "Frio y Fx Ambientales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   165
         TabIndex        =   43
         Top             =   300
         Width           =   2580
         Begin VB.TextBox txtCmdExec 
            Height          =   285
            Left            =   1320
            TabIndex        =   53
            Text            =   "0"
            Top             =   1110
            Width           =   915
         End
         Begin VB.TextBox txtIntervaloPerdidaStaminaLluvia 
            Height          =   300
            Left            =   1320
            TabIndex        =   51
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloWAVFX 
            Height          =   300
            Left            =   150
            TabIndex        =   45
            Text            =   "0"
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloFrio 
            Height          =   285
            Left            =   180
            TabIndex        =   44
            Text            =   "0"
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "TimerExec"
            Height          =   195
            Left            =   1320
            TabIndex        =   54
            Top             =   840
            Width           =   750
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Stamina Lluvia"
            Height          =   195
            Left            =   1350
            TabIndex        =   52
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "FxS"
            Height          =   195
            Left            =   180
            TabIndex        =   47
            Top             =   270
            Width           =   270
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Frio"
            Height          =   195
            Left            =   195
            TabIndex        =   46
            Top             =   810
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Non Player Characters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   2775
      TabIndex        =   36
      Top             =   2160
      Width           =   2655
      Begin VB.Frame Frame4 
         Caption         =   "A.I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   150
         TabIndex        =   37
         Top             =   240
         Width           =   1365
         Begin VB.TextBox txtAI 
            Height          =   285
            Left            =   150
            TabIndex        =   39
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.TextBox txtNPCPuedeAtacar 
            Height          =   285
            Left            =   135
            TabIndex        =   38
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "AI"
            Height          =   195
            Left            =   165
            TabIndex        =   41
            Top             =   930
            Width           =   150
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Puede atacar"
            Height          =   195
            Left            =   150
            TabIndex        =   40
            Top             =   255
            Width           =   960
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   105
      TabIndex        =   3
      Top             =   45
      Width           =   7455
      Begin VB.Frame Frame9 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   90
         TabIndex        =   24
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtIntervaloParaConexion 
            Height          =   300
            Left            =   45
            TabIndex        =   26
            Text            =   "0"
            Top             =   495
            Width           =   930
         End
         Begin VB.TextBox txtTrabajo 
            Height          =   300
            Left            =   60
            TabIndex        =   25
            Text            =   "0"
            Top             =   1140
            Width           =   930
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "IntervaloCon"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Trabajo"
            Height          =   195
            Left            =   165
            TabIndex        =   27
            Top             =   900
            Width           =   540
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Combate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   1545
         TabIndex        =   19
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtPuedeAtacar 
            Height          =   300
            Left            =   135
            TabIndex        =   22
            Text            =   "0"
            Top             =   1200
            Width           =   930
         End
         Begin VB.TextBox txtIntervaloLanzaHechizo 
            Height          =   300
            Left            =   150
            TabIndex        =   20
            Text            =   "0"
            Top             =   525
            Width           =   930
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Puede Atacar"
            Height          =   195
            Left            =   135
            TabIndex        =   23
            Top             =   930
            Width           =   975
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Lanza Spell"
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   285
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Hambre y sed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   5925
         TabIndex        =   14
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtIntervaloHambre 
            Height          =   285
            Left            =   150
            TabIndex        =   16
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtIntervaloSed 
            Height          =   285
            Left            =   150
            TabIndex        =   15
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hambre"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   255
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sed"
            Height          =   195
            Left            =   165
            TabIndex        =   17
            Top             =   930
            Width           =   285
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sanar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   4470
         TabIndex        =   9
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtSanaIntervaloDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   11
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtSanaIntervaloSinDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   10
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descansando"
            Height          =   195
            Left            =   180
            TabIndex        =   13
            Top             =   255
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sin descansar"
            Height          =   195
            Left            =   165
            TabIndex        =   12
            Top             =   930
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stamina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   3015
         TabIndex        =   4
         Top             =   210
         Width           =   1410
         Begin VB.TextBox txtStaminaIntervaloSinDescansar 
            Height          =   285
            Left            =   150
            TabIndex        =   6
            Text            =   "0"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.TextBox txtStaminaIntervaloDescansar 
            Height          =   285
            Left            =   165
            TabIndex        =   5
            Text            =   "0"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sin descansar"
            Height          =   195
            Left            =   165
            TabIndex        =   8
            Top             =   930
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descansando"
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   255
            Width           =   990
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Magia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   2130
      Width           =   7455
      Begin VB.Frame Frame10 
         Caption         =   "Duracion Spells"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   135
         TabIndex        =   29
         Top             =   270
         Width           =   2400
         Begin VB.TextBox txtInvocacion 
            Height          =   300
            Left            =   1170
            TabIndex        =   49
            Text            =   "0"
            Top             =   1170
            Width           =   900
         End
         Begin VB.TextBox txtIntervaloInvisible 
            Height          =   300
            Left            =   1170
            TabIndex        =   34
            Text            =   "0"
            Top             =   495
            Width           =   900
         End
         Begin VB.TextBox txtIntervaloParalizado 
            Height          =   300
            Left            =   195
            TabIndex        =   31
            Text            =   "0"
            Top             =   1170
            Width           =   795
         End
         Begin VB.TextBox txtIntervaloVeneno 
            Height          =   300
            Left            =   195
            TabIndex        =   30
            Text            =   "0"
            Top             =   510
            Width           =   795
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Invocacion"
            Height          =   195
            Left            =   1170
            TabIndex        =   50
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Invisible"
            Height          =   195
            Left            =   1170
            TabIndex        =   35
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Paralizado"
            Height          =   195
            Left            =   225
            TabIndex        =   33
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Veneno"
            Height          =   180
            Left            =   225
            TabIndex        =   32
            Top             =   300
            Width           =   555
         End
      End
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   6480
      Width           =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   0
      Top             =   6480
      Width           =   2040
   End
   Begin VB.Label Label22 
      Caption         =   "Regeneración Vampiro"
      Height          =   255
      Left            =   1680
      TabIndex        =   58
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label21 
      Caption         =   "Intervalo Flechas"
      Height          =   255
      Left            =   120
      TabIndex        =   55
      Top             =   4680
      Width           =   1335
   End
End
Attribute VB_Name = "FrmInterv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub AplicarIntervalos()

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿ Intervalos del main loop ¿?¿?¿?¿?¿?¿?¿?¿?¿
    SanaIntervaloSinDescansar = val(txtSanaIntervaloSinDescansar.Text)
    StaminaIntervaloSinDescansar = val(txtStaminaIntervaloSinDescansar.Text)
    SanaIntervaloDescansar = val(txtSanaIntervaloDescansar.Text)
    StaminaIntervaloDescansar = val(txtStaminaIntervaloDescansar.Text)
    IntervaloSed = val(txtIntervaloSed.Text)
    IntervaloHambre = val(txtIntervaloHambre.Text)
    IntervaloVeneno = val(txtIntervaloVeneno.Text)
    IntervaloParalizado = val(txtIntervaloParalizado.Text)
    IntervaloInvisible = val(txtIntervaloInvisible.Text)
    IntervaloFrio = val(txtIntervaloFrio.Text)
    IntervaloWavFx = val(txtIntervaloWAVFX.Text)
    IntervaloInvocacion = val(txtInvocacion.Text)
    IntervaloParaConexion = val(txtIntervaloParaConexion.Text)

    '///////////////// TIMERS \\\\\\\\\\\\\\\\\\\

    IntervaloUserPuedeCastear = val(txtIntervaloLanzaHechizo.Text)
    frmMain.npcataca.Interval = val(txtNPCPuedeAtacar.Text)
    frmMain.TIMER_AI.Interval = val(txtAI.Text)
    IntervaloUserPuedeTrabajar = val(txtTrabajo.Text)
    'pluto:2.8.0
    IntervaloUserPuedeFlechas = val(TxtFlechas.Text)
    IntervaloRegeneraVampiro = val(txtVampire.Text)

    IntervaloUserPuedeAtacar = val(txtPuedeAtacar.Text)
    frmMain.tLluvia.Interval = val(txtIntervaloPerdidaStaminaLluvia.Text)
    frmMain.CmdExec.Interval = val(txtCmdExec.Text)


End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Call AplicarIntervalos

End Sub

Private Sub Command2_Click()

    On Error GoTo Err

    'Intervalos
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar", str(SanaIntervaloSinDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar", str(StaminaIntervaloSinDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar", str(SanaIntervaloDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar", str(StaminaIntervaloDescansar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed", str(IntervaloSed))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre", str(IntervaloHambre))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno", str(IntervaloVeneno))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado", str(IntervaloParalizado))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible", str(IntervaloInvisible))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio", str(IntervaloFrio))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX", str(IntervaloWavFx))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion", str(IntervaloInvocacion))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion", str(IntervaloParaConexion))

    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    'pluto:2.8.0
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeFlechas", str(IntervaloUserPuedeFlechas))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloRegeneraVampiro", str(IntervaloRegeneraVampiro))

    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo", str(IntervaloUserPuedeCastear))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI", frmMain.TIMER_AI.Interval)
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar", frmMain.npcataca.Interval)
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo", str(IntervaloUserPuedeTrabajar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar", str(IntervaloUserPuedeAtacar))
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPerdidaStaminaLluvia", frmMain.tLluvia.Interval)
    Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTimerExec", frmMain.CmdExec.Interval)

    MsgBox "Los intervalos se han guardado sin problemas"

    Exit Sub
Err:
    MsgBox "Error al intentar grabar los intervalos"
End Sub

Private Sub ok_Click()
    Me.Visible = False
End Sub

Private Sub Text1_Change()

End Sub

