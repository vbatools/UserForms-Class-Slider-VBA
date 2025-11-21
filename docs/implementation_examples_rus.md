# Примеры реализации для clsSlider

## Содержание
1. [Базовая реализация](#базовая-реализация)
2. [Расширенная настройка](#расширенная-настройка)
3. [Пользовательский стиль](#пользовательский-стиль)
4. [Обработка событий](#обработка-событий)
5. [Интеграция с формами](#интеграция-с-формами)
6. [Динамическое создание слайдеров](#динамическое-создание-слайдеров)
7. [Работа с несколькими слайдерами](#работа-с-несколькими-слайдерами)
8. [Продвинутые примеры](#продвинутые-примеры)

## Базовая реализация

### Простое использование слайдера
```vba
' В модуле формы
Dim Slider As clsSlider

Private Sub UserForm_Initialize()
    ' Создание экземпляра класса слайдера
    Set Slider = New clsSlider
    
    ' Инициализация слайдера с параметрами по умолчанию
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True)
End Sub
```

### Слайдер с отображением значения
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    
    ' Инициализация с видимой меткой значения
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True, "0.0")
    
    ' Обновление метки с текущим значением
    LabelValue.Caption = Slider.Value
End Sub

Private Sub Slider_ValueChanged()
    ' Обновление при изменении значения слайдера
    LabelValue.Caption = Slider.Value
End Sub
```

## Расширенная настройка

### Настройка цветов и внешнего вида
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    
    ' Инициализация с пользовательскими цветами
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True, "0.0", clsSlider.PositionValue.Right, _
        RGB(200, 200, 200), RGB(0, 100, 200), RGB(0, 0, 0), RGB(0, 0, 0))
        
    ' Дальнейшая настройка внешнего вида
    With Slider
        .ForeColorValue = RGB(0, 100, 200)
        .BackColorFull = RGB(0, 150, 255)
        .BackColorEmpty = RGB(220, 220, 220)
    End With
End Sub
```

### Горизонтальный и вертикальный слайдер
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    
    ' Горизонтальный слайдер (по умолчанию) - шире чем высокий
    Call Slider.Initialize(Me.HorizontalLabel, 50, 0, 100, True)
    
    ' Вертикальный слайдер - выше чем широкий
    Call Slider.Initialize(Me.VerticalLabel, 50, 0, 100, True)
    
    ' Проверка ориентации
    If Slider.IsHorizontal Then
        Debug.Print "Слайдер горизонтальный"
    Else
        Debug.Print "Слайдер вертикальный"
    End If
End Sub
```

## Пользовательский стиль

### Пользовательский формат значения
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    
    ' Инициализация с пользовательским форматированием значения
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True, "0.00%", clsSlider.PositionValue.Top)
    
    ' Это будет отображать значения в процентах (например, 50%)
End Sub
```

### Слайдер с иконкой
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    
    ' Инициализация с пользовательской иконкой на кнопке
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True, , , , , , , 59963) ' Использование кода иконки
End Sub
```

## Обработка событий

### Обработка изменений значения
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True)
End Sub

Private Sub mLabelBtn_Click()
    ' Обработка события клика
    Debug.Print "Слайдер нажат, текущее значение: " & Slider.Value
End Sub

Private Sub mLabelBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Обработка движения мыши (перетаскивание)
    Debug.Print "Значение слайдера изменилось на: " & Slider.Value
    ' Обновление других элементов управления на основе значения слайдера
    TextBox1.Value = Slider.Value
End Sub
```

### Блокировка слайдера
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True)
    
    ' Изначально заблокировать слайдер
    Slider.Locked = True
End Sub

Private Sub LockSliderButton_Click()
    ' Переключение состояния блокировки слайдера
    Slider.Locked = Not Slider.Locked
    If Slider.Locked Then
        LockButton.Caption = "Разблокировать слайдер"
    Else
        LockButton.Caption = "Заблокировать слайдер"
    End If
End Sub
```

## Интеграция с формами

### Полный пример формы
```vba
Dim VolumeSlider As clsSlider
Dim BrightnessSlider As clsSlider

Private Sub UserForm_Initialize()
    ' Инициализация слайдера громкости
    Set VolumeSlider = New clsSlider
    Call VolumeSlider.Initialize(Me.VolumeLabel, 75, 0, 100, True, "0", clsSlider.PositionValue.Right)
    
    ' Инициализация слайдера яркости
    Set BrightnessSlider = New clsSlider
    Call BrightnessSlider.Initialize(Me.BrightnessLabel, 50, 0, 100, True, "0", clsSlider.PositionValue.Right)
    
    ' Обновление элементов управления формы начальными значениями
    UpdateVolumeDisplay
    UpdateBrightnessDisplay
End Sub

Private Sub UpdateVolumeDisplay()
    VolumeValueLabel.Caption = "Громкость: " & VolumeSlider.Value & "%"
    ' Симуляция эффекта изменения громкости
    ' (В реальном приложении это бы управляло громкостью аудио)
End Sub

Private Sub UpdateBrightnessDisplay()
    BrightnessValueLabel.Caption = "Яркость: " & BrightnessSlider.Value & "%"
    ' Симуляция эффекта изменения яркости
    ' (В реальном приложении это бы управляло яркостью экрана)
End Sub

Private Sub VolumeLabel_ValueChanged()
    UpdateVolumeDisplay
End Sub

Private Sub BrightnessLabel_ValueChanged()
    UpdateBrightnessDisplay
End Sub
```

## Динамическое создание слайдеров

### Создание слайдеров программно
```vba
Private Sub UserForm_Initialize()
    ' Создание нескольких слайдеров динамически
    Dim i As Integer
    Dim newLabel As MSForms.Label
    Dim newSlider As clsSlider
    
    For i = 1 To 3
        ' Создание новой метки для слайдера
        Set newLabel = Me.Controls.Add("Forms.Label.1", "DynamicSlider" & i, True)
        With newLabel
            .Left = 20
            .Top = 30 + (i - 1) * 40
            .Width = 200
            .Height = 5
            .BackColor = RGB(200, 200, 200)
        End With
        
        ' Создание и инициализация слайдера
        Set newSlider = New clsSlider
        Call newSlider.Initialize(newLabel, 50, 0, 100, True, , clsSlider.PositionValue.Bottom)
        
        ' Сохранение ссылки в коллекции для последующего доступа
        SliderCollection.Add newSlider
    Next i
End Sub
```

## Работа с несколькими слайдерами

### Управление несколькими слайдерами
```vba
Dim SliderCollection As Collection

Private Sub UserForm_Initialize()
    Set SliderCollection = New Collection
    
    ' Инициализация нескольких слайдеров
    InitializeSliders
End Sub

Private Sub InitializeSliders()
    Dim RedSlider As clsSlider
    Dim GreenSlider As clsSlider
    Dim BlueSlider As clsSlider
    
    ' Создание RGB слайдеров
    Set RedSlider = New clsSlider
    Set GreenSlider = New clsSlider
    Set BlueSlider = New clsSlider
    
    ' Инициализация с разными параметрами
    Call RedSlider.Initialize(Me.RedLabel, 128, 0, 255, True, "0", clsSlider.PositionValue.Right)
    Call GreenSlider.Initialize(Me.GreenLabel, 128, 0, 255, True, "0", clsSlider.PositionValue.Right)
    Call BlueSlider.Initialize(Me.BlueLabel, 128, 0, 255, True, "0", clsSlider.PositionValue.Right)
    
    ' Добавление в коллекцию
    SliderCollection.Add RedSlider
    SliderCollection.Add GreenSlider
    SliderCollection.Add BlueSlider
    
    ' Обновление предварительного просмотра цвета
    UpdateColorPreview
End Sub

Private Sub UpdateColorPreview()
    ' Объединение значений слайдеров для создания цвета
    Dim combinedColor As Long
    combinedColor = RGB(SliderCollection(1).Value, SliderCollection(2).Value, SliderCollection(3).Value)
    
    ' Применение цвета к области предварительного просмотра
    ColorPreview.BackColor = combinedColor
    ColorPreview.Caption = "RGB(" & SliderCollection(1).Value & "," & _
                           SliderCollection(2).Value & "," & _
                           SliderCollection(3).Value & ")"
End Sub

Private Sub RedLabel_ValueChanged()
    UpdateColorPreview
End Sub

Private Sub GreenLabel_ValueChanged()
    UpdateColorPreview
End Sub

Private Sub BlueLabel_ValueChanged()
    UpdateColorPreview
End Sub
```

## Продвинутые примеры

### Слайдер с эффектом анимации
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True)
End Sub

Private Sub AnimateSliderEffect()
    Dim originalColor As Long
    originalColor = Slider.BackColorFull
    
    ' Создание эффекта анимации путем временного изменения цвета
    Slider.BackColorFull = RGB(255, 255, 0) ' Желтый
    
    ' Краткое ожидание для отображения эффекта
    Application.Wait (Now + TimeValue("0:00:0.2"))
    
    ' Восстановление исходного цвета
    Slider.BackColorFull = originalColor
End Sub

Private Sub mLabelBtn_Click()
    ' Запуск анимации при клике
    AnimateSliderEffect
End Sub
```

### Адаптивная панель слайдеров
```vba
Private Sub UserForm_Resize()
    ' Настройка позиций слайдеров при изменении размера формы
    AdjustSliderLayout
End Sub

Private Sub AdjustSliderLayout()
    If Not Slider Is Nothing Then
        With Slider
            ' Настройка слайдера для сохранения пропорций
            .control.Width = Me.Width - 60
            .control.Left = 30
            
            ' Если слайдер вертикальный, настройка высоты
            If Not .IsHorizontal Then
                .control.Height = Me.Height - 100
                .control.Top = 50
            End If
        End With
    End If
End Sub
```

Эти примеры демонстрируют различные способы реализации и использования класса `clsSlider` в VBA UserForms. Каждый пример может быть адаптирован под конкретные требования приложения.