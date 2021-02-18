VERSION 5.00
Begin VB.Form frmPrincipal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reindexar dados"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8715
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtZip 
      Height          =   330
      Left            =   2550
      TabIndex        =   1
      Text            =   "Esta TextBox é usado internamente pelo AddZip, podendo fica escondida do usuário."
      Top             =   2700
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6780
      Top             =   1920
   End
   Begin VB.Label lblinfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Aguardando conexão com o banco de dados...."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   3075
      Left            =   -2160
      Picture         =   "Form1.frx":0000
      Top             =   -30
      Width           =   8850
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cn As Object
Dim Instancia As String
Private Sub Label1_Click()

End Sub
'---------------------------------------------------------------------------------------
' Procedure : Logs
' DateTime  : 09/09/2019 01:42
' Author    : Soares
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Logs(valor As String)
   On Error GoTo Logs_Error

    'Open App.Path & "\" & Instancia & "-" & Format(Now(), " DD-MM-YYYY") & ".txt" For Append As #1
    'Print #1, valor
    'Close #1

   On Error GoTo 0
   Exit Function

Logs_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Logs of Formulário frmPrincipal na linha " & Erl
End Function
'---------------------------------------------------------------------------------------
' Procedure : InicioProcesso
' DateTime  : 09/09/2019 01:38
' Author    : Soares
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub InicioProcesso()
 '   On Error GoTo Erro
    
   On Error GoTo InicioProcesso_Error

    Timer1.Enabled = False

    Dim valor() As String
    If (Command = "") Then
        MsgBox "Paramentros incorretos!", vbCritical, "Atenção"
        End
    End If
    valor = Split(Command, "!")

    Dim stringCn As String
    stringCn = "Provider=SQLOLEDB.1;Password=%3;Persist Security Info=True;User ID=%2;Initial Catalog=master;Data Source=%1"
    stringCn = Replace(stringCn, "%1", valor(0))
    stringCn = Replace(stringCn, "%2", valor(1))
    stringCn = Replace(stringCn, "%3", valor(2))



    Instancia = Replace(valor(0), "\", "")
    Instancia = Replace(Instancia, "-", "")
    Instancia = Replace(Instancia, "/", "")

    lblinfo.Caption = "Aguardando conexão com o banco de dados [" & valor(0) & "]"
    DoEvents
    Set Cn = CreateObject("Adodb.Connection")
    Cn.Open stringCn

    Logs "Inicio da Execução - " & Now()
    Dim RsDataBases As Object
    Dim RsTable As Object
    Dim ini As String
    Dim Fim As String

    Set RsDataBases = Cn.Execute("select name From sysdatabases")
    'where left(name, 2)='ng'")

    Dim caminho As String
    
    'If UBound(valor) > 5 Then
    '    caminho = valor(6)
    '    If Right(caminho, 1) <> "\" Then caminho = caminho & "\"
    '    caminho = caminho & Instancia & "-" & Format(Now, "ddmmyyyy-hhss")
    '
    'Else
        caminho = "E:\BackupMegaNG\SQL\" & Instancia & "-" & Format(Now, "ddmmyyyy-hhss") & ""
    'End If
    On Error Resume Next
    'MsgBox "Antes Criei a pasta"
    Call MkDir(caminho)
    'MsgBox "Depois Criei a pasta"
    On Error GoTo InicioProcesso_Error

    Do While Not RsDataBases.EOF
        
        If LCase(Left(RsDataBases.fields("name").Value, 2)) = "ng" Then
            'If (UCase(valor(3)) = "SIM") Then
                Set RsTable = Cn.Execute("select name From " & RsDataBases.fields("name") & ".dbo.sysobjects where xtype = 'U'")
                Logs "Selecionando os banco de dados - " & RsDataBases.fields("name")
                Do While Not RsTable.EOF
                    ini = "" & Time
                    lblinfo.Caption = "Reindexando tabela " & valor(0) & " --> " & RsDataBases.fields("name") & "." & RsTable.fields("name").Value
                    DoEvents
                    Call retRecordSet("dbcc dbreindex('" & RsDataBases.fields("name") & ".dbo." & RsTable.fields("name").Value & "')")
                    Fim = "" & Time
                    Logs "Reindex tabala " & RsTable.fields("name").Value & " Inicio : " & ini & " Fim " & Fim
                    RsTable.MoveNext
                    DoEvents
                Loop
            'End If
            'If (UCase(valor(4)) = "SIM") Then
                lblinfo.Caption = "SHRINKDATABASE " & RsDataBases.fields("name")
                DoEvents
                Logs "DBCC SHRINKDATABASE(N'" & RsDataBases.fields("name") & ")"
                DoEvents
                Call retRecordSet("DBCC SHRINKDATABASE(N'" & RsDataBases.fields("name") & "')")
            'End If
            'If (UCase(valor(5)) = "SIM") Then
                lblinfo.Caption = "BACKUP " & RsDataBases.fields("name")
                DoEvents
                Logs "Fazendo backup " & RsDataBases.fields("name") & ""
                DoEvents
                sql = "BACKUP DATABASE [" & RsDataBases.fields("name") & "] TO  DISK = N'" & caminho & "\" & RsDataBases.fields("name") & ".bak' WITH NOFORMAT, NOINIT,  NAME = N'" & RsDataBases.fields("name") & "-Cheio Banco de Dados Backup', Skip , NOREWIND, NOUNLOAD, STATS = 10"
                'Open App.Path & "\exe.sql" For Append As #99
                'Print #99, sql
                'Close #99
                Call retRecordSet(sql)
                lblinfo.Caption = "Gerando Zip - " & RsDataBases.fields("name")
                DoEvents
                
                Logs "Gerando zip do arquivo - " & RsDataBases.fields("name") & ""
                DoEvents
                
                Compacta caminho & "\" & RsDataBases.fields("name") & ".zip", caminho & "\" & RsDataBases.fields("name") & ".bak"
                DoEvents
                If Dir(caminho & "\" & RsDataBases.fields("name") & ".zip", vbArchive) <> "" Then
                    On Error Resume Next
                    Kill caminho & "\" & RsDataBases.fields("name") & ".bak"
                    On Error GoTo InicioProcesso_Error
                End If
                DoEvents
            'End If
        End If
        RsDataBases.MoveNext
        DoEvents
    Loop
    Logs "Fim da Execução - " & Now()

    End

'Erro:
    'Logs "Erro: " & Err.Number & " - " & Err.Description
    End

   On Error GoTo 0
   Exit Sub

InicioProcesso_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InicioProcesso of Formulário frmPrincipal na linha " & Erl
End
End Sub


Public Function retRecordSet(StrSQL)
On Error GoTo Erro1
    Dim cmd ' as new ADODB.Command
    Dim rs 'As New ADODB.Recordset
    
    Set cmd = CreateObject("ADODB.Command")
    Set rs = CreateObject("ADODB.Recordset")
    
    cmd.ActiveConnection = Cn
    cmd.CommandText = StrSQL
    cmd.CommandTimeout = 0
    Set rs = cmd.Execute
    
    Set retRecordSet = rs
    Exit Function
    
Erro1:
    Logs "Erro: " & Err.Number & " - " & Err.Description
End Function

Private Sub Form_Load()
    InicializaZip Me, txtZip
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call InicioProcesso
    Timer1.Enabled = True
End Sub
