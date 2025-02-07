VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Дурак"
   ClientHeight    =   8820
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   10965
   Icon            =   "Igra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Управление"
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   8040
      Width           =   7335
      Begin VB.CommandButton Command3 
         Caption         =   "Бито"
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Взять..."
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Прозрачно
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Прозрачно
         Caption         =   "Ходит:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   11040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Новая Игра"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7320
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image zakrper 
      Height          =   1065
      Left            =   8760
      Picture         =   "Igra.frx":406A
      Top             =   2400
      Width           =   1440
   End
   Begin VB.Image iotvet 
      Height          =   375
      Index           =   0
      Left            =   5520
      Top             =   8280
      Width           =   375
   End
   Begin VB.Image prom 
      Height          =   375
      Left            =   8760
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lub 
      BackStyle       =   0  'Прозрачно
      Caption         =   "0"
      Height          =   255
      Left            =   9600
      TabIndex        =   14
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label lui 
      BackStyle       =   0  'Прозрачно
      Caption         =   "0"
      Height          =   255
      Left            =   9600
      TabIndex        =   13
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label luk 
      BackStyle       =   0  'Прозрачно
      Caption         =   "0"
      Height          =   255
      Left            =   9600
      TabIndex        =   12
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lukl 
      BackStyle       =   0  'Прозрачно
      Caption         =   "36"
      Height          =   255
      Left            =   9600
      TabIndex        =   11
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Бито:"
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Прозрачно
      Caption         =   "У игрока:"
      Height          =   255
      Left            =   8160
      TabIndex        =   9
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Прозрачно
      Caption         =   "У компьютера:"
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Карт в колоде:"
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Image komp 
      Height          =   375
      Index           =   0
      Left            =   4920
      Top             =   8400
      Width           =   495
   End
   Begin VB.Image Igrok 
      Height          =   375
      Index           =   0
      Left            =   4320
      Top             =   8400
      Width           =   495
   End
   Begin VB.Image kzkart 
      Height          =   1455
      Left            =   9000
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image dano 
      Height          =   495
      Index           =   0
      Left            =   6120
      Top             =   8280
      Width           =   495
   End
   Begin VB.Menu fail 
      Caption         =   "Файл"
      Begin VB.Menu new 
         Caption         =   "Новая игра"
      End
      Begin VB.Menu Exit 
         Caption         =   "Выход "
      End
   End
   Begin VB.Menu help 
      Caption         =   "О программе"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m(1 To 36, 1 To 7) As Integer ' Массив карт с их статусом
'Dim n(1 To 36) As Integer
Dim igraidet As Boolean
Dim h As Integer ' используется для присвоения значения random
Dim kzs As String ' козырная карта тип String
Dim okis As String
Dim kzi As Integer ' козырная карта тип Integer
Dim uk As Integer ' Количество карт у компьютера
Dim ui As Integer ' играка
Dim ub As Integer ' в бито
Dim ukl As Integer ' в колоде
Dim ud As Integer
Dim ukr As Integer
Dim hod As Integer ' чей ход
Dim otvet As Integer
Dim km As Integer ' Для определении козырной масти
Dim nki As Integer
Dim nkk As Integer
Dim minkartb As Boolean ' для работы с картами
Dim Vzat As Boolean
Dim minkart As Integer

Private Sub Command1_Click()
NewGames
End Sub

Private Sub NewGames()
Timer1.Enabled = True
ProvStat
Och1
Peremesh
Koz
PerRazd
RashKart
KtoHodit
OtobrKartIg
OtobrKartKom
End Sub

Private Sub ProvStat() ' перед загруской проверяем играли раньше или нет
If igraidet Then
For i = 1 To 13
On Error GoTo e1
Unload Igrok(i)
Next i
e1:
two
tree
four
End If
igraidet = True
End Sub

Private Sub two()
For j = 1 To 13
On Error GoTo e2
Unload komp(j)
Next j
e2:
End Sub
Private Sub tree()
For l = 1 To 30
On Error GoTo e3
Unload dano(l)
Next l
e3:
End Sub

Private Sub four()
For po = 1 To 30
On Error GoTo e4
Unload iotvet(po)
Next po
e4:
End Sub

Private Sub Command2_Click()
If hod = 2 Or hod = 3 Then
If otvet = 1 Then
ubratKD
ubratKR
For i = 1 To 36
m(i, 6) = 5
If (m(i, 3) = 5) Or (m(i, 3) = 6) Then
m(i, 3) = 1
End If
Next i
UbratKI
RashKart
OtobrKartIg
CheiHotOtobr
IzKolodiKomp
IzKolodiIg
hod = 3
otvet = 2
End If
End If
CheiHotOtobr
End Sub

Private Sub Command3_Click()
If (hod = 1) Or (hod = 4) Then
If otvet = 1 Then
ubratKD
ubratKR
For i = 1 To 36
m(i, 6) = 5
    If (m(i, 3) = 5) Or (m(i, 3) = 6) Then
    m(i, 3) = 4
    RashKart
    End If
Next i
End If
hod = 3
otvet = 2
CheiHotOtobr
IzKolodiKomp
IzKolodiIg

End If

End Sub

Private Sub Command4_Click()
If Vzat = False Then
If (hod = 1) Or (hod = 4) Then
hod = 3
otvet = 2
Else
hod = 1
otvet = 1
End If
End If
If Vzat Then
hod = 1
otvet = 1
Vzat = False
IzKolodiKomp
IzKolodiIg

End If

Command4.Enabled = False
Command3.Enabled = True
End Sub
Private Sub six()
If Vzat = False Then
If (hod = 1) Or (hod = 4) Then
hod = 3
otvet = 2
Else
hod = 1
otvet = 1
End If
End If
If Vzat Then
hod = 1
otvet = 1
Vzat = False
IzKolodiKomp
IzKolodiIg
End If
End Sub
Private Sub IzKolodiIg()
If ukl <> 0 Then
If ui < 6 Then

za4:
If ukl <> 0 Then
For o = 1 To 36
If (m(o, 3) = 3) Then
UbratKI
m(o, 3) = 1
PereShet
RashKart
OtobrKartIg
End If

If ui = 6 Then GoTo za3
Next o
If ui < 6 Then GoTo za4
za3:
End If
End If
End If
End Sub
Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
igraidet = False
otvet = 2
End Sub

'1-значеник, 2-масть, 3-где находятся
'4-расположение в колоде
Private Sub help_Click()
Spravka.Show
End Sub

Private Sub Peremesh()
'Процедура для перемешивания карт в колоде
k = 5
l = 0
For i = 1 To 9
k = k + 1
l = l + 1
m(l, 2) = 1 ' Черви
m(l, 1) = k ' Значение карты начиная с 6
m(l, 3) = 3 ' расположение карты с начала все в колоде (3)
m(l, 5) = l ' порядковый номер
m(l, 6) = 3 ' вспомогательный
l = l + 1
m(l, 2) = 2 ' Буби
m(l, 1) = k
m(l, 3) = 3
m(l, 5) = l
m(l, 6) = 3
l = l + 1
m(l, 2) = 3 ' Пики
m(l, 1) = k
m(l, 3) = 3
m(l, 5) = l
m(l, 6) = 3
l = l + 1
m(l, 2) = 4 ' Крести
m(l, 1) = k
m(l, 3) = 3
m(l, 5) = l
m(l, 6) = 3
Next i
p = False
Randomize
For i = 1 To 36
m1: h = Rnd * 10 * 4 - 4 ' присваиваем случайное число
If h <= 0 Then GoTo m1 ' проверяем если число отрицательное или равно нулю отправляем назад
For j = 1 To 36
If m(j, 4) = h Then GoTo m1 ' проверяем есть ли в колоде карта с таким же раположением если есть то отпровляем назад
Next j
m(i, 4) = h ' если все прошло хорошо присваиваем текущей карте расположение
Next i
End Sub

Private Sub Koz() ' Определяем козырную карту
kzs = ""
For i = 1 To 36
If m(i, 4) = 36 Then kzi = m(i, 5) ' Ищем последнию карту она будет козырная
Next i
If m(kzi, 2) = 1 Then
kzs = "\karti\chervi"
km = 1
End If
If m(kzi, 2) = 2 Then
kzs = "\karti\bubi"
km = 2
End If
If m(kzi, 2) = 3 Then
kzs = "\karti\piki"
km = 3
End If
If m(kzi, 2) = 4 Then
kzs = "\karti\kresti"
km = 4
End If
Text1.Text = m(kzi, 1)
kzs = kzs + Text1.Text
kzs = kzs + ".bmp"
kzkart.Picture = LoadPicture(App.Path + kzs)
End Sub

Private Sub Och1() ' Очищаем значения карт
For i = 1 To 36
m(i, 4) = 0
Next i
End Sub

Private Sub PerRazd() ' первая раздача карт
m2:
For o = 1 To 36
If (m(o, 4) = 1) And (m(o, 3) = 3) Then
m(o, 3) = 2
PereShet
End If
RashKart
If uk = 6 Then GoTo m3
Next o
If uk < 6 Then GoTo m2
m3:

For o = 1 To 36
If (m(o, 4) = 1) And (m(o, 3) = 3) Then
m(o, 3) = 1
PereShet
End If
RashKart
If ui = 6 Then GoTo m4
Next o
If ui < 6 Then GoTo m3

m4:
End Sub

Private Sub PereShet() ' перещитываем карты смещаем в колоде на одну выше
For Y = 1 To 36
If m(Y, 3) = 3 Then m(Y, 4) = m(Y, 4) - 1
Next Y
End Sub

Private Sub RashKart() ' ращитываем сколько карт где находятся
ui = 0
uk = 0
ub = 0
ukl = 0
ud = 0
ukr = 0
For e = 1 To 36
    If m(e, 3) = 1 Then ui = ui + 1 ' у игрока
    If m(e, 3) = 2 Then uk = uk + 1 ' компа
    If m(e, 3) = 3 Then ukl = ukl + 1 ' в колоде
    If m(e, 3) = 4 Then ub = ub + 1 ' в бито
    If m(e, 3) = 5 Then ud = ud + 1 ' в дано
    If m(e, 3) = 6 Then ukr = ukr + 1 ' крыто
Next e
lukl.Caption = ukl
luk.Caption = uk
lui.Caption = ui
lub.Caption = ub
End Sub

Private Sub KtoHodit() ' определяем кто будет ходить первым
nki = 36
nkk = 36
For i = 1 To 36
    If (m(i, 3) = 1) And (m(i, 1) < nki) And (m(i, 2) = km) Then nki = m(i, 1)
    If (m(i, 3) = 2) And (m(i, 1) < nkk) And (m(i, 2) = km) Then nkk = m(i, 1)
Next i

If nki < nkk Then
hod = 1
otvet = 1
Else
hod = 3
otvet = 2
End If
CheiHotOtobr
End Sub

Private Sub CheiHotOtobr()
If hod = 2 Then Label2.Caption = "Компьютер"
If hod = 3 Then Label2.Caption = "Компьютер"
If hod = 1 Then Label2.Caption = "Игрок"
If hod = 4 Then Label2.Caption = "Игрок"
End Sub

Private Sub OtobrKartIg() ' отображаем карты игрока
i = 1
For j = 1 To 36
    If m(j, 3) = 1 Then
    Load Igrok(i)
    OKI (j)
    Igrok(i).Picture = LoadPicture(App.Path + okis)
    Igrok(i).Tag = m(j, 5)
    Igrok(i).Top = Height - 3000
    Igrok(i).Left = ((Form1.Width / (ui + 1)) * i) - 600
    Igrok(i).Visible = True
    i = i + 1
    End If
Next j
End Sub

Private Sub Igrok_Click(Index As Integer)
minkart = 675

If (hod = 1) And (otvet = 1) Then
    minkart = m(Igrok(Index).Tag, 5)
    VyitKartIgroka2
    otvet = 2
    hod = 4
End If

If (hod = 4) And (otvet = 1) Then
For i = 1 To 36
    If (m(i, 3) = 5) And (m(i, 1) = m(Igrok(Index).Tag, 1)) Then
    minkart = m(Igrok(Index).Tag, 5)
    VyitKartIgroka2
    otvet = 2
    GoTo we1
    End If
Next i
End If

If (hod = 4) And (otvet = 1) Then
    For i = 1 To 36
        If (m(i, 3) = 6) And (m(i, 1) = m(Igrok(Index).Tag, 1)) Then
        minkart = m(Igrok(Index).Tag, 5)
        VyitKartIgroka2
        otvet = 2
        GoTo we1
        End If
    Next i
    End If
we1:

    If (hod = 2) And (otvet = 1) Then
    For i = 1 To 36
        If (m(i, 3) = 5) And (m(i, 6) <> 1) And (m(i, 1) < m(Igrok(Index).Tag, 1)) And (m(i, 2) = m(Igrok(Index).Tag, 2)) Then
        minkart = Igrok(Index).Tag
        VyitKartIgroka
        m(i, 6) = 1
        otvet = 2
        GoTo we2
        End If
        If (m(i, 3) = 5) And (m(i, 6) <> 1) And (m(i, 2) <> km) And (m(Igrok(Index).Tag, 2) = km) Then
        minkart = Igrok(Index).Tag
        VyitKartIgroka
        m(i, 6) = 1
        otvet = 2
        GoTo we2
        End If
    Next i
we2:
End If

End Sub
Private Sub VyitKartIgroka()
For i = 1 To ui
    If minkart = Igrok(i).Tag Then
    prom.Picture = Igrok(i).Picture
    prom.Tag = Igrok(i).Tag
    m(minkart, 3) = 6
    UbratKI
    RashKart
    Load iotvet(ukr)
    iotvet(ukr).Visible = True
    iotvet(ukr).Top = 3850
    iotvet(ukr).Left = (ud * 1100) - 800
    iotvet(ukr).Picture = prom.Picture
    iotvet(ukr).Tag = prom.Tag
    OtobrKartIg
    GoTo r2
    End If
Next i
r2:
End Sub

Private Sub VyitKartIgroka2() ' вытягиваем карты игрока komp - компа соответстванно
For i = 1 To ui
    If minkart = Igrok(i).Tag Then
    prom.Picture = Igrok(i).Picture
    prom.Tag = Igrok(i).Tag
    m(minkart, 3) = 5
    UbratKI
    RashKart
    Load dano(ud)
    dano(ud).Visible = True
    dano(ud).Top = 2400
    dano(ud).Left = (ud * 1100) - 800
    dano(ud).Picture = prom.Picture
    dano(ud).Tag = prom.Tag
    OtobrKartIg
    GoTo r2
    End If
Next i
r2:
End Sub

Private Sub UbratKI() ' убираем карты игрока
For ad = 1 To ui
Unload Igrok(ad)
Next ad
End Sub


Private Sub OKI(qw) ' присваив переменной значения пути

If m(qw, 2) = 1 Then
okis = "\karti\chervi"

End If
If m(qw, 2) = 2 Then
okis = "\karti\bubi"

End If
If m(qw, 2) = 3 Then
okis = "\karti\piki"

End If
If m(qw, 2) = 4 Then
okis = "\karti\kresti"

End If
Text2.Text = m(qw, 1)
okis = okis + Text2.Text
okis = okis + ".bmp"
End Sub

Private Sub OtobrKartKom() ' отоброжаем карты компа igrok-соответственно
i = 1
For j = 1 To 36
    If m(j, 3) = 2 Then
    Load komp(i)
    OKI (j)
    komp(i).Picture = LoadPicture(App.Path + "\karti\SNOWMEN.BMP")
    komp(i).Tag = m(j, 5)
    komp(i).Top = 600
    komp(i).Left = ((Form1.Width / (uk + 1)) * i) - 600
    komp(i).Visible = True
    i = i + 1
    End If
Next j
End Sub

Private Sub new_Click()
NewGames
End Sub

Private Sub Timer1_Timer()
minkartb = False
Pobeda
If (hod = 2) And (otvet = 2) Then HodKompa
If (hod = 3) And (otvet = 2) Then HodKompa
If (hod = 1) And (otvet = 2) Then HodKompa
If (hod = 4) And (otvet = 2) Then HodKompa
End Sub

Private Sub Pobeda() ' определ кто выиграл
If ud = 0 And ui = 0 Then
qwer = MsgBox("Игрок выиграл!")
ubratKR
ubratKD
UbratKI
UbratKK
NewGames
End If
If ud = 0 And uk = 0 Then
qwer = MsgBox("Компьютер выиграл!")
ubratKR
ubratKD
UbratKI
UbratKK
NewGames
End If
If ui > 14 Then
qwer = MsgBox("Компьютер выиграл!")
ubratKR
ubratKD
UbratKI
UbratKK
NewGames

End If
If uk > 14 Then
qwer = MsgBox("Игрок выиграл!")
ubratKR
ubratKD
UbratKI
UbratKK
NewGames
End If
End Sub

Private Sub HodKompa() ' это все ужасное действия компа в определенных случаях (мозг)
minkart = 675
If (hod = 2) Or (hod = 3) Then
    For i = 1 To 36
        If (hod = 3) And (ui <> 0) Then
        If (m(i, 3) = 2) And (m(i, 2) <> km) Then
        minkartb = True
        If minkart > m(i, 5) Then
        minkart = m(i, 5)
        End If
        End If
        End If
    Next i
hod = 2
otvet = 1

    If (hod = 3) And (minkartb = False) And (ui <> 0) Then
        For i = 1 To 36
            If (m(i, 3) = 2) Then
            If minkart > m(i, 5) Then
            minkart = m(i, 5)
            otvet = 1
            GoTo wq1
            End If
            End If
        Next i
    End If

If minkartb Then GoTo wq1

    For i = 1 To 36
        If (hod = 2) And (m(i, 3) = 2) And (ui <> 0) Then
        For j = 1 To 36
            If (m(i, 1) = m(j, 1)) And (m(j, 3) = 5) Then
            minkart = m(i, 5)
            otvet = 1
            GoTo wq1
            End If
        Next j
        For j = 1 To 36
            If (m(i, 1) = m(j, 1)) And (m(j, 3) = 6) Then
            minkart = m(i, 5)
            otvet = 1
            GoTo wq1
            End If
        Next j
        End If
    Next i

    
    
    otvet = 1
    BitoCom
wq1:     VyitKartKompa
GoTo qwe
End If

If (hod = 1) Or (hod = 4) Then
    For i = 1 To 36
    If (m(i, 3) = 5) And (m(i, 6) <> 1) Then
        For j = 1 To 36
            If (m(j, 3) = 2) And (m(i, 1) < m(j, 1)) And (m(i, 2) = m(j, 2)) Then
            minkart = m(j, 5)
            VyitKartKompa2
            m(i, 6) = 1
            otvet = 1
            GoTo qwe
            End If
        Next j
    End If
    Next i

    For i = 1 To 36
    If (m(i, 3) = 5) And (m(i, 6) <> 1) Then
        For j = 1 To 36
            If (m(i, 2) <> km) And (m(j, 2) = km) And (m(j, 3) = 2) Then
            minkart = m(j, 5)
            VyitKartKompa2
            m(i, 6) = 1
            otvet = 1
            GoTo qwe
            End If
        Next j
        End If
    Next i
    VzatCom
    otvet = 1
End If
qwe:
End Sub

Private Sub VzatCom() ' если комп подымает
ubratKD
ubratKR
For i = 1 To 36
m(i, 6) = 5
If (m(i, 3) = 5) Or (m(i, 3) = 6) Then
m(i, 3) = 2
End If
Next i
UbratKK
RashKart
OtobrKartKom
CheiHotOtobr
IzKolodiIg
Vzat = True
six
CheiHotOtobr
End Sub

Private Sub BitoCom() ' если комп кладет в бито
ubratKD
ubratKR
For i = 1 To 36
m(i, 6) = 5
    If (m(i, 3) = 5) Or (m(i, 3) = 6) Then
    m(i, 3) = 4
    RashKart
    End If
Next i
IzKolodiKomp
IzKolodiIg
six
CheiHotOtobr
End Sub

Private Sub IzKolodiKomp() 'берем карты из колоды
If ukl <> 0 Then
If uk < 6 Then

za2:
If ukl <> 0 Then
For o = 1 To 36
If (m(o, 3) = 3) Then
UbratKK
m(o, 3) = 2
PereShet
RashKart
OtobrKartKom
End If

If uk = 6 Then GoTo za1
Next o
If uk < 6 Then GoTo za2
za1:
End If
End If
End If
End Sub

Private Sub ubratKD() ' убир. карты которыми ходили
For i = 1 To ud
Unload dano(i)
Next i
End Sub

Private Sub ubratKR() ' убир. карты которыми отбивались
For i = 1 To ukr
Unload iotvet(i)
Next i
End Sub

Private Sub VyitKartKompa2()
For i = 1 To uk
    If minkart = komp(i).Tag Then 'берем минимальную карту у компа
    OKI (minkart)
    prom.Picture = LoadPicture(App.Path + okis)
    prom.Tag = komp(i).Tag
    m(minkart, 3) = 6
    UbratKK
    RashKart
    Load iotvet(ukr)
    iotvet(ukr).Visible = True
    iotvet(ukr).Top = 3850
    iotvet(ukr).Left = (ukr * 1100) - 800
    iotvet(ukr).Picture = prom.Picture ' и помещаем на стол
    iotvet(ukr).Tag = prom.Tag
    PokazKK
    GoTo r5
    End If
Next i
r5:
End Sub

Private Sub VyitKartKompa()
For i = 1 To uk
    If minkart = komp(i).Tag Then 'берем минимальную карту у компа
    OKI (minkart)
    prom.Picture = LoadPicture(App.Path + okis)
    prom.Tag = komp(i).Tag
    m(minkart, 3) = 5
    UbratKK
    RashKart
    Load dano(ud)
    dano(ud).Visible = True
    dano(ud).Top = 2400
    dano(ud).Left = (ud * 1100) - 800
    dano(ud).Picture = prom.Picture ' и помещаем на стол
    dano(ud).Tag = prom.Tag
    PokazKK
    GoTo r1
    End If
Next i
r1:

End Sub
Private Sub one()

End Sub

Private Sub PokazKK() ' показываем карты компа
i = 1
For j = 1 To 36
    If m(j, 3) = 2 Then
    Load komp(i)
    OKI (j)
    komp(i).Picture = LoadPicture(App.Path + "\karti\SNOWMEN.BMP")
    komp(i).Tag = m(j, 5)
    komp(i).Top = 600
    komp(i).Left = ((Form1.Width / (uk + 1)) * i) - 600
    komp(i).Visible = True
    i = i + 1
    End If
Next j
End Sub

Private Sub UbratKK() ' убир. карты компа
For ad = 1 To uk
    Unload komp(ad)
Next ad
End Sub

