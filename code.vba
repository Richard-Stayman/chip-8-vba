Option Explicit

' Константы

Const Millisecond As Double = 0.000000011574

Dim PowerOfTwo(15) As Long

' Переменные

Private Type ScreenPixel
    value As Byte
    Dirty As Byte
End Type

Private Type VM
    ' Состояние виртуальной машины
    Running As Boolean

    ' Регистры
    PC As Long ' Программный счётчик
    i As Long ' Регистр I, 16 бит, но Integer не подойдёт т.к. со знаком
    V(15) As Byte ' Регистры V0 по VF

    ' Стек вызовов
    Stack(15) As Long ' Адреса
    StackTop As Long ' верхушка стека

    ' Память
    Memory(4095) As Byte

    ' Экран
    Screen(63, 31) As ScreenPixel

    ' Таймер задержки
    DelayTimer As Byte
    SoundTimer As Byte

    ' Текущая нажатая клавиша, -1 если ничего не нажато
    CurrentKey As Integer
End Type

Dim ThisVM As VM
Dim temp As Integer

Dim LastRedraw As Double

Private Sub cb0_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 0
End Sub

Private Sub cb0_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cb1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 1
End Sub

Private Sub cb1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cb2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 2
End Sub

Private Sub cb2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cb3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 3
End Sub

Private Sub cb3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cb4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 4
End Sub

Private Sub cb4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cb5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 5
End Sub

Private Sub cb5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cb6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 6
End Sub

Private Sub cb6_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cb7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 7
End Sub

Private Sub cb7_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cb8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 8
End Sub

Private Sub cb8_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cb9_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 9
End Sub

Private Sub cb9_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cbA_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 10
End Sub

Private Sub cbA_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cbB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 11
End Sub

Private Sub cbB_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cbC_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 12
End Sub

Private Sub cbC_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cbD_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 13
End Sub

Private Sub cbD_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cbE_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 14
End Sub

Private Sub cbE_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cbF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = 15
End Sub

Private Sub cbF_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ThisVM.CurrentKey = -1
End Sub

Private Sub cbClearScreen_Click()
    ClearScreen
    Redraw
End Sub


Public Sub SetScreenPixel(ByVal X As Long, ByVal Y As Long, ByVal value As Byte)
    With ThisVM.Screen(X, Y)
        .value = value
        .Dirty = 1
    End With
End Sub

Public Sub ClearScreen()
    Dim i, j As Long
    For i = 0 To 63
        For j = 0 To 31
            SetScreenPixel i, j, 0
        Next j
    Next i

    Redraw
End Sub

Private Sub InitConstants()
    Dim i, lastN As Long
    lastN = 1
    PowerOfTwo(0) = lastN
    For i = 1 To 15
        lastN = lastN * 2
        PowerOfTwo(i) = lastN
    Next i
End Sub

Private Sub InitVM()
    InitConstants

    Randomize

    ThisVM.CurrentKey = -1

    ThisVM.PC = 512
    ThisVM.i = 0
    Dim i As Long
    For i = 0 To 15
        ThisVM.V(i) = 0
    Next i
    ThisVM.DelayTimer = 0
    ThisVM.SoundTimer = 0

    ThisVM.Memory(0) = &HF0&
    ThisVM.Memory(1) = &H90&
    ThisVM.Memory(2) = &H90&
    ThisVM.Memory(3) = &H90&
    ThisVM.Memory(4) = &HF0&
    ThisVM.Memory(5) = &H20&
    ThisVM.Memory(6) = &H60&
    ThisVM.Memory(7) = &H20&
    ThisVM.Memory(8) = &H20&
    ThisVM.Memory(9) = &H70&
    ThisVM.Memory(10) = &HF0&
    ThisVM.Memory(11) = &H10&
    ThisVM.Memory(12) = &HF0&
    ThisVM.Memory(13) = &H80&
    ThisVM.Memory(14) = &HF0&
    ThisVM.Memory(15) = &HF0&
    ThisVM.Memory(16) = &H10&
    ThisVM.Memory(17) = &HF0&
    ThisVM.Memory(18) = &H10&
    ThisVM.Memory(19) = &HF0&
    ThisVM.Memory(20) = &H90&
    ThisVM.Memory(21) = &H90&
    ThisVM.Memory(22) = &HF0&
    ThisVM.Memory(23) = &H10&
    ThisVM.Memory(24) = &H10&
    ThisVM.Memory(25) = &HF0&
    ThisVM.Memory(26) = &H80&
    ThisVM.Memory(27) = &HF0&
    ThisVM.Memory(28) = &H10&
    ThisVM.Memory(29) = &HF0&
    ThisVM.Memory(30) = &HF0&
    ThisVM.Memory(31) = &H80&
    ThisVM.Memory(32) = &HF0&
    ThisVM.Memory(33) = &H90&
    ThisVM.Memory(34) = &HF0&
    ThisVM.Memory(35) = &HF0&
    ThisVM.Memory(36) = &H10&
    ThisVM.Memory(37) = &H20&
    ThisVM.Memory(38) = &H40&
    ThisVM.Memory(39) = &H40&
    ThisVM.Memory(40) = &HF0&
    ThisVM.Memory(41) = &H90&
    ThisVM.Memory(42) = &HF0&
    ThisVM.Memory(43) = &H90&
    ThisVM.Memory(44) = &HF0&
    ThisVM.Memory(45) = &HF0&
    ThisVM.Memory(46) = &H90&
    ThisVM.Memory(47) = &HF0&
    ThisVM.Memory(48) = &H10&
    ThisVM.Memory(49) = &HF0&
    ThisVM.Memory(50) = &HF0&
    ThisVM.Memory(51) = &H90&
    ThisVM.Memory(52) = &HF0&
    ThisVM.Memory(53) = &H90&
    ThisVM.Memory(54) = &H90&
    ThisVM.Memory(55) = &HE0&
    ThisVM.Memory(56) = &H90&
    ThisVM.Memory(57) = &HE0&
    ThisVM.Memory(58) = &H90&
    ThisVM.Memory(59) = &HE0&
    ThisVM.Memory(60) = &HF0&
    ThisVM.Memory(61) = &H80&
    ThisVM.Memory(62) = &H80&
    ThisVM.Memory(63) = &H80&
    ThisVM.Memory(64) = &HF0&
    ThisVM.Memory(65) = &HE0&
    ThisVM.Memory(66) = &H90&
    ThisVM.Memory(67) = &H90&
    ThisVM.Memory(68) = &H90&
    ThisVM.Memory(69) = &HE0&
    ThisVM.Memory(70) = &HF0&
    ThisVM.Memory(71) = &H80&
    ThisVM.Memory(72) = &HF0&
    ThisVM.Memory(73) = &H80&
    ThisVM.Memory(74) = &HF0&
    ThisVM.Memory(75) = &HF0&
    ThisVM.Memory(76) = &H80&
    ThisVM.Memory(77) = &HF0&
    ThisVM.Memory(78) = &H80&
    ThisVM.Memory(79) = &H80&

    ThisVM.StackTop = 0

    ClearScreen

    VMSetRunning False
End Sub

Private Sub cbStep_Click()
    Step
    Redraw
End Sub

Private Sub CommandButton1_Click()
    Dim filePath As Variant
    Dim rom() As Byte
    Dim file As Integer: file = FreeFile
    Dim i As Long

    filePath = Application.GetOpenFilename
    If filePath = False Then
        MsgBox "Загрузка картриджа отменена."
        Exit Sub
    End If

    Open filePath For Binary Access Read As #file
    ReDim rom(0 To LOF(file) - 1)
    Get #file, , rom
    Close #file

    For i = 0 To UBound(rom)
        ThisVM.Memory(i + 512) = rom(i)
    Next i

    InitVM
End Sub

Private Sub VMSetRunning(Running As Boolean)
    ThisVM.Running = Running

    If ThisVM.Running Then
        CommandButton2.Caption = "Стоп"
    Else
        CommandButton2.Caption = "Старт"
    End If
End Sub

Private Sub CommandButton2_Click()
    VMSetRunning (Not ThisVM.Running)

    If ThisVM.Running Then
        LastRedraw = Now
        ProcessVM
    End If
End Sub

Private Function VMGetShort() As Long
    Dim i As Long
    ' Преобразуем byte в long для операции сдвига
    i = (CLng(ThisVM.Memory(ThisVM.PC)) * 256) Or CLng(ThisVM.Memory(ThisVM.PC + 1))
    ThisVM.PC = ThisVM.PC + 2

    VMGetShort = i
End Function

Private Sub TODO(str As String)
    MsgBox "TODO: " & str & " (Адрес: " & Hex(ThisVM.PC) & ")"
End Sub

' Спрайты состоят из полосок шириной в 8 и произвольной высотой до 15
' Если спрайт выходит за пределы экрана, то он должен вылезти с другой стороны
Private Function DrawSprite(ByVal address As Long, ByVal targetX As Long, ByVal targetY As Long, ByVal bytes As Long) As Byte
    If bytes = 0 Then
        Exit Function
    End If

    Dim a As Long
    Dim spriteByte As Byte
    Dim j, Y As Long
    Dim i, X As Long
    Dim bit As Boolean
    Dim bitMask As Long

    DrawSprite = 0

    j = targetY
    For a = address To address + bytes - 1
        spriteByte = ThisVM.Memory(a)

        Y = j Mod 32
        bitMask = 1
        For i = targetX + 7 To targetX Step -1
            X = i Mod 64

            bit = spriteByte And bitMask
            If bit Then
                If ThisVM.Screen(X, Y).value <> 0 Then
                    DrawSprite = 1
                    SetScreenPixel X, Y, 0
                Else
                    SetScreenPixel X, Y, 1
                End If
            End If

            bitMask = bitMask * 2
        Next i

        j = j + 1
    Next a
End Function

Public Sub Step()
    ' Получаем инструкцию
    '''' ВНИМАНИЕ!!! Для шеснадцатеричных чисел больше 32 тысяч запись выглядит не как &H1234, а как &H1234&

    ' Всяческие переменные, используемые в Select Case
    Dim register, value, registerAnd, X, Y, n, i As Long

    Dim instruction As Long: instruction = VMGetShort
    Select Case (instruction And &HF000&) ' Маска для выделения одного четырёхбитного числа
        Case &H0& ' Инструкция системного вызова
            Dim syscall As Integer: syscall = CInt(instruction And &HFFF)  ' Маска для выделения 12-битного адреса системного вызова
            Select Case syscall
                Case &HE0  ' Очистка экрана
                    ClearScreen
                Case &HEE  ' Возврат из процедуры
                    If ThisVM.StackTop = 0 Then
                        TODO "Недополнение стека"
                        VMSetRunning False
                    Else
                        ThisVM.StackTop = ThisVM.StackTop - 1
                        ThisVM.PC = ThisVM.Stack(ThisVM.StackTop)
                    End If
            End Select

        Case &H1000& ' Инструкция прыжка по адресу
            ThisVM.PC = instruction And &HFFF

        Case &H2000& ' Инструкция вызова процедуры
            If ThisVM.StackTop = 16 Then
                TODO "Переполнение стека"
                VMSetRunning False
            Else
                ThisVM.Stack(ThisVM.StackTop) = ThisVM.PC ' В PC уже хранится указатель на следующую инструкцию
                ThisVM.StackTop = ThisVM.StackTop + 1
                ThisVM.PC = instruction And &HFFF
            End If

        Case &H3000& ' Инструкция пропуска следующей инструкции, если регистр равен параметру
            register = (instruction And &HF00&) \ PowerOfTwo(8) ' получаем 4 битный регистр
            value = (instruction And &HFF&) ' получаем 8 битный параметр

            If ThisVM.V(register) = value Then
                ThisVM.PC = ThisVM.PC + 2
            End If

        Case &H4000& ' Инструкция пропуска следующей инструкции, если регистр НЕ равен параметру
            register = (instruction And &HF00&) \ PowerOfTwo(8) ' получаем 4 битный регистр
            value = (instruction And &HFF&) ' получаем 8 битный параметр

            If ThisVM.V(register) <> value Then
                ThisVM.PC = ThisVM.PC + 2
            End If

        Case &H5000& ' Инструкция пропуска следующей инструкции при сравнении регистра с регистром
            X = (instruction And &HF00&) \ PowerOfTwo(8) ' получаем 4 битный регистр
            Y = (instruction And &HF0&) \ PowerOfTwo(4)
            n = (instruction And &HF&)

            Select Case n
                Case 0
                    If ThisVM.V(X) = ThisVM.V(Y) Then
                        ThisVM.PC = ThisVM.PC + 2
                    End If

                Case Else
                    TODO "Нет операции условного перехода " & Hex(n) & " с регистрами V" & Hex(X) & " и V" & Hex(Y)
                    VMSetRunning False
            End Select

        Case &H6000& ' Инструкция приравнивания значения регистра к аргументу
            register = (instruction And &HF00&) \ PowerOfTwo(8) ' получаем 4 битный регистр
            value = (instruction And &HFF&) ' получаем 8 битный параметр
            ThisVM.V(register) = value

        Case &H7000& ' Инструкция прибавления параметра к регистру, игнорируя переполнением
            register = (instruction And &HF00&) \ PowerOfTwo(8)
            value = (instruction And &HFF&) ' получаем 8 битный параметр
            ThisVM.V(register) = CByte((CInt(ThisVM.V(register)) + value) Mod 256)

        Case &H8000& ' Инструкция с различными операциями для двух регистров
            X = (instruction And &HF00&) \ PowerOfTwo(8)
            Y = (instruction And &HF0&) \ PowerOfTwo(4)
            n = (instruction And &HF&)

            Select Case n
                Case &H0& ' Загрузить значение Vy в Vx
                    ThisVM.V(X) = ThisVM.V(Y)

                Case &H1& ' Побитовое ИЛИ
                    ThisVM.V(X) = ThisVM.V(X) Or ThisVM.V(Y)
                    ThisVM.V(&HF&) = 0

                Case &H2& ' Побитовое И
                    ThisVM.V(X) = ThisVM.V(X) And ThisVM.V(Y)
                    ThisVM.V(&HF&) = 0

                Case &H3& ' Побитовое сложение по модулю 2
                    ThisVM.V(X) = ThisVM.V(X) Xor ThisVM.V(Y)
                    ThisVM.V(&HF&) = 0

                Case &H4& ' Сложение с флагом переполнения
                    value = CLng(ThisVM.V(X)) + CLng(ThisVM.V(Y))
                    If value > 255 Then
                        temp = 1
                    Else
                        ThisVM.V(&HF&) = 0
                    End If
                    ThisVM.V(X) = value Mod 256
                    ThisVM.V(&HF&) = temp

                Case &H5& ' Вычесть Vy из Vx. Если Vx > Vy, то VF = 1, иначе 0
                    If ThisVM.V(X) >= ThisVM.V(Y) Then
                        ThisVM.V(X) = ThisVM.V(X) - ThisVM.V(Y)
                        temp = 1
                    Else
                        ' При переполнении вычитаем меньшее из большего, затем результат из 256
                        ThisVM.V(X) = (&H100& - (ThisVM.V(Y) - ThisVM.V(X))) And &HFF&
                        temp = 0
                    End If
                    ThisVM.V(&HF&) = temp

                Case &H6& ' Побитовый сдвиг направо, VF равен младшему знаку числа в регистре Vx
                    temp = ThisVM.V(X) And &H1&
                    ThisVM.V(X) = ThisVM.V(X) \ 2
                    ThisVM.V(&HF&) = temp

                Case &H7& '
                If ThisVM.V(Y) >= ThisVM.V(X) Then
                    temp = 1
                Else
                    temp = 0
                End If
                If ThisVM.V(Y) >= ThisVM.V(X) Then
                    ThisVM.V(X) = ThisVM.V(Y) - ThisVM.V(X)
                Else
                    ThisVM.V(X) = (&H100& - (ThisVM.V(X) - ThisVM.V(Y))) And &HFF&
                End If
                ThisVM.V(&HF&) = temp

                Case &HE& ' Побитовый сдвиг налево, VF равен старшему знаку числа в регистре Vx
                    temp = ThisVM.V(X) \ PowerOfTwo(7)
                    ThisVM.V(X) = (ThisVM.V(X) And &H7F&) * 2 ' Маска, исключающая самый левый бит во избежание переполнения
                    ThisVM.V(&HF&) = temp

                Case Else
                    TODO "Нет операции " & Hex(n) & " над регистрами V" & Hex(X) & " и V" & Hex(Y)
                    VMSetRunning False
            End Select

        Case &H9000& ' Ещё одна инструкция пропуска следующей инструкции при сравнении регистра с регистром
            X = (instruction And &HF00&) \ PowerOfTwo(8) ' получаем 4 битный регистр
            Y = (instruction And &HF0&) \ PowerOfTwo(4)
            n = (instruction And &HF&)

            Select Case n
                Case 0
                    If ThisVM.V(X) <> ThisVM.V(Y) Then
                        ThisVM.PC = ThisVM.PC + 2
                    End If

                Case Else
                    TODO "Нет операции условного перехода " & Hex(n) & " с регистрами V" & Hex(X) & " и V" & Hex(Y)
                    VMSetRunning False
            End Select

        Case &HA000& ' Инструкция загрузки числа в регистр I
            ThisVM.i = instruction And &HFFF ' Маска для выделения 12-битного числа из инструкции

        Case &HB000&
            ThisVM.PC = instruction + ThisVM.V(0) And &HFFF

        Case &HC000& ' Инструкция загрузки случайного числа в регистр
            register = (instruction And &HF00&) \ PowerOfTwo(8) ' получаем 4 битный регистр
            registerAnd = (instruction And &HFF&) ' получаем 8 битный параметр
            ThisVM.V(register) = Int(PowerOfTwo(8) * Rnd) And registerAnd ' Случайное число в [0; 256)

        Case &HD000& ' Инструкция отрисовки спрайта на экран
            X = ThisVM.V((instruction And &HF00&) \ PowerOfTwo(8))
            Y = ThisVM.V((instruction And &HF0&) \ PowerOfTwo(4))
            n = (instruction And &HF&)
            ThisVM.V(&HF) = DrawSprite(ThisVM.i, X, Y, n)

        Case &HE000& ' Инструкция условного перехода по клавише
            X = ThisVM.V((instruction And &HF00&) \ PowerOfTwo(8))
            n = (instruction And &HFF&)

            Select Case n
                Case &H9E& ' Пропустить следующую инструкцию при нажатой клавише [Vx]
                    If ThisVM.V(X) = ThisVM.CurrentKey Then
                        ThisVM.PC = ThisVM.PC + 2
                    End If

                Case &HA1& ' Пропустить следующую инструкцию при ненажатой клавише [Vx]
                    If ThisVM.V(X) <> ThisVM.CurrentKey Then
                        ThisVM.PC = ThisVM.PC + 2
                    End If

                Case Else
                    TODO "Нет операции условного пропуска " & Hex(n) & " с регистром V" & Hex(X)
                    VMSetRunning False
            End Select

        Case &HF000& ' Прочие инструкции
            X = (instruction And &HF00&) \ PowerOfTwo(8)
            n = (instruction And &HFF&)

            Select Case n
                Case &H7& ' Запись таймера в регистр Vx
                    ThisVM.V(X) = ThisVM.DelayTimer

                Case &HA& ' Ожидание нажатия клавиши
                    Redraw
                    While ThisVM.CurrentKey = -1
                        DoEvents
                    Wend
                    ThisVM.V(X) = ThisVM.CurrentKey

                Case &H15& ' Задание таймера задержки
                    ThisVM.DelayTimer = ThisVM.V(X)

                Case &H18& ' Задание таймера звука
                    ThisVM.SoundTimer = ThisVM.V(X)

                Case &H1E& ' Прибавление Vx к I
                    ThisVM.i = ThisVM.i + ThisVM.V(X)

                Case &H29& ' Получение адреса спрайта цифры из регистра Vx
                    ThisVM.i = ThisVM.V(X) * 5

                Case &H33& ' Преобразование числа из регистра Vx в десятичную форму по адресам I, I+1 и I+2
                    ThisVM.Memory(ThisVM.i) = ThisVM.V(X) \ 100
                    ThisVM.Memory(ThisVM.i + 1) = (ThisVM.V(X) \ 10) Mod 10
                    ThisVM.Memory(ThisVM.i + 2) = ThisVM.V(X) Mod 10

                Case &H55& ' Сохранить все регистры по адресу I
                    For i = 0 To X
                        ThisVM.Memory(ThisVM.i + i) = ThisVM.V(i)
                    Next i

                Case &H65& ' Загрузить все регистры по адресу I
                    For i = 0 To X
                        ThisVM.V(i) = ThisVM.Memory(ThisVM.i + i)
                    Next i

                Case Else
                    TODO "Нет операции " & Hex(n) & " над регистром V" & Hex(X)
                    VMSetRunning False
            End Select

        Case Else
            TODO "Непонятная инструкция " & Hex(instruction)
            ThisVM.PC = ThisVM.PC - 2 ' Откатываемся назад ровно на одну инструкцию
            VMSetRunning False
    End Select
End Sub

Public Sub ProcessVM()
    If Not ThisVM.Running Then
        Exit Sub
    End If

    Step

    ' Работаем как можно быстрей
    Dim Count
    Count = Now + Millisecond

    If LastRedraw + Millisecond * 16 < Now Then
        Redraw

        ' Обновление таймера задержки
        If ThisVM.DelayTimer > 0 Then
            ThisVM.DelayTimer = ThisVM.DelayTimer - 1
        End If

        ' Обновление таймера звука
        If ThisVM.SoundTimer > 0 Then
            ' TODO: проиграть звук
            ThisVM.SoundTimer = ThisVM.SoundTimer - 1
        End If
    End If

    DoEvents
    Application.OnTime Count, "CHIP8List.ProcessVM"
End Sub

Private Sub RedrawDebug()
    Range("$BO$3").value = Hex(ThisVM.PC)
    Range("$BO$4").value = Hex(ThisVM.i)
    Range("$BO$5").value = Hex(ThisVM.DelayTimer)

    Dim i As Long
    For i = 0 To 15
        Range("$BR$" & (i + 3)).value = Hex(ThisVM.V(i))
    Next i
End Sub


Public Sub Redraw()
    LastRedraw = Now

    RedrawDebug

    Dim i, j As Long
    For i = 0 To 63
        For j = 0 To 31
            If ThisVM.Screen(i, j).Dirty = 1 Then
                ThisVM.Screen(i, j).Dirty = 0
                If ThisVM.Screen(i, j).value = 1 Then
                    Cells(j + 1, i + 1).Interior.ColorIndex = 2 ' WHITE
                Else
                    Cells(j + 1, i + 1).Interior.ColorIndex = 1 ' BLACK
                End If
            End If
        Next j
    Next i
End Sub
