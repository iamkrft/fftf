Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Создаем главную форму
$form = New-Object System.Windows.Forms.Form
$form.Text = "Вывод слов с интервалом"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Элементы главной формы
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20, 20)
$label.Size = New-Object System.Drawing.Size(360, 30)
$label.Text = "Выберите файл со словами:"
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(20, 50)
$textBox.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($textBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(280, 50)
$browseButton.Size = New-Object System.Drawing.Size(75, 23)
$browseButton.Text = "Обзор..."
$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $textBox.Text = $openFileDialog.FileName
    }
})
$form.Controls.Add($browseButton)

$intervalLabel = New-Object System.Windows.Forms.Label
$intervalLabel.Location = New-Object System.Drawing.Point(20, 90)
$intervalLabel.Size = New-Object System.Drawing.Size(200, 20)
$intervalLabel.Text = "Интервал (секунды):"
$form.Controls.Add($intervalLabel)

$intervalBox = New-Object System.Windows.Forms.NumericUpDown
$intervalBox.Location = New-Object System.Drawing.Point(20, 110)
$intervalBox.Size = New-Object System.Drawing.Size(100, 20)
$intervalBox.Minimum = 1
$intervalBox.Maximum = 3600
$intervalBox.Value = 60
$form.Controls.Add($intervalBox)

# Чекбокс для перемешивания слов
$shuffleCheckbox = New-Object System.Windows.Forms.CheckBox
$shuffleCheckbox.Location = New-Object System.Drawing.Point(20, 140)
$shuffleCheckbox.Size = New-Object System.Drawing.Size(200, 20)
$shuffleCheckbox.Text = "Перемешать слова"
$shuffleCheckbox.Checked = $false
$form.Controls.Add($shuffleCheckbox)

# Настройки цвета
$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Location = New-Object System.Drawing.Point(20, 170)
$colorLabel.Size = New-Object System.Drawing.Size(200, 20)
$colorLabel.Text = "Цвет слов:"
$form.Controls.Add($colorLabel)

# Выбор цвета
$colorComboBox = New-Object System.Windows.Forms.ComboBox
$colorComboBox.Location = New-Object System.Drawing.Point(20, 190)
$colorComboBox.Size = New-Object System.Drawing.Size(150, 20)
$colorComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
# Добавляем цвета
[Enum]::GetValues([System.Drawing.KnownColor]) | Where-Object { 
    $_ -notmatch "Active|Inactive|Button|Control|Desktop|Gradient|GrayText|Highlight|HotTrack|Menu|ScrollBar|Window"
} | ForEach-Object {
    $colorComboBox.Items.Add($_) | Out-Null
}
$colorComboBox.SelectedItem = "Black"
$form.Controls.Add($colorComboBox)

# Чекбокс для случайного цвета
$randomColorCheckbox = New-Object System.Windows.Forms.CheckBox
$randomColorCheckbox.Location = New-Object System.Drawing.Point(180, 190)
$randomColorCheckbox.Size = New-Object System.Drawing.Size(200, 20)
$randomColorCheckbox.Text = "Случайный цвет"
$randomColorCheckbox.Checked = $false
$form.Controls.Add($randomColorCheckbox)

$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(20, 230)
$startButton.Size = New-Object System.Drawing.Size(100, 30)
$startButton.Text = "Старт"
$startButton.Add_Click({
    if (-not [string]::IsNullOrEmpty($textBox.Text)) {
        if (Test-Path $textBox.Text) {
            $words = Get-Content $textBox.Text
            if ($words.Count -gt 0) {
                
                # Перемешиваем слова если нужно
                if ($shuffleCheckbox.Checked) {
                    $words = $words | Sort-Object {Get-Random}
                }
                
                $form.Hide()
                
                # Создаем форму для отображения слов
                $displayForm = New-Object System.Windows.Forms.Form
                $displayForm.Text = "Текущее слово"
                $displayForm.Size = New-Object System.Drawing.Size(500, 300)
                $displayForm.StartPosition = "CenterScreen"
                $displayForm.FormBorderStyle = "FixedDialog"
                $displayForm.MaximizeBox = $false
                
                $wordLabel = New-Object System.Windows.Forms.Label
                $wordLabel.Font = New-Object System.Drawing.Font("Arial", 24, [System.Drawing.FontStyle]::Bold)
                $wordLabel.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
                $wordLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
                $displayForm.Controls.Add($wordLabel)
                
                # Панель для кнопок
                $buttonPanel = New-Object System.Windows.Forms.Panel
                $buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
                $buttonPanel.Height = 40
                $displayForm.Controls.Add($buttonPanel)
                
                # Кнопка "Пропустить"
                $skipButton = New-Object System.Windows.Forms.Button
                $skipButton.Text = "Пропустить"
                $skipButton.Size = New-Object System.Drawing.Size(100, 30)
                $skipButton.Location = New-Object System.Drawing.Point(20, 5)
                $skipButton.Add_Click({
                    $script:currentIndex++
                    $timer.Stop()
                    ShowNextWord
                    $timer.Start()
                })
                $buttonPanel.Controls.Add($skipButton)
                
                # Кнопка "Стоп"
                $stopButton = New-Object System.Windows.Forms.Button
                $stopButton.Text = "Стоп"
                $stopButton.Size = New-Object System.Drawing.Size(100, 30)
                $stopButton.Location = New-Object System.Drawing.Point(140, 5)
                $stopButton.Add_Click({
                    $displayForm.Close()
                    $form.Show()
                })
                $buttonPanel.Controls.Add($stopButton)
                
                # Кнопка "Перемешать"
                $reshuffleButton = New-Object System.Windows.Forms.Button
                $reshuffleButton.Text = "Перемешать"
                $reshuffleButton.Size = New-Object System.Drawing.Size(100, 30)
                $reshuffleButton.Location = New-Object System.Drawing.Point(260, 5)
                $reshuffleButton.Add_Click({
                    # Сохраняем уже показанные слова
                    $shownWords = $words[0..($script:currentIndex-1)]
                    $remainingWords = $words[$script:currentIndex..($words.Count-1)]
                    
                    # Перемешиваем оставшиеся слова
                    $remainingWords = $remainingWords | Sort-Object {Get-Random}
                    
                    # Собираем обратно
                    $script:words = $shownWords + $remainingWords
                    $script:currentIndex = $shownWords.Count
                    
                    [System.Windows.Forms.MessageBox]::Show("Оставшиеся слова были перемешаны!", "Информация")
                })
                $buttonPanel.Controls.Add($reshuffleButton)
                
                $displayForm.Add_FormClosing({
                    $timer.Stop()
                    $timer.Dispose()
                })
                
                $timer = New-Object System.Windows.Forms.Timer
                $timer.Interval = $intervalBox.Value * 1000
                $script:currentIndex = 0
                $script:words = $words
                $script:selectedColor = $colorComboBox.SelectedItem
                $script:useRandomColor = $randomColorCheckbox.Checked
                
                # Функция для получения случайного цвета
                function Get-RandomColor {
                    $colors = [Enum]::GetValues([System.Drawing.KnownColor]) | Where-Object { 
                        $_ -notmatch "Active|Inactive|Button|Control|Desktop|Gradient|GrayText|Highlight|HotTrack|Menu|ScrollBar|Window"
                    }
                    return [System.Drawing.Color]::FromKnownColor(($colors | Get-Random))
                }
                
                # Функция для показа следующего слова
                function ShowNextWord {
                    if ($script:currentIndex -lt $script:words.Count) {
                        $wordLabel.Text = $script:words[$script:currentIndex]
                        $displayForm.Text = "Текущее слово ($($script:currentIndex+1)/$($script:words.Count))"
                        
                        # Устанавливаем цвет слова
                        if ($script:useRandomColor) {
                            $wordLabel.ForeColor = Get-RandomColor
                        } else {
                            $wordLabel.ForeColor = [System.Drawing.Color]::FromKnownColor($script:selectedColor)
                        }
                    } else {
                        $timer.Stop()
                        [System.Windows.Forms.MessageBox]::Show("Все слова были показаны!", "Завершено")
                        $displayForm.Close()
                        $form.Show()
                    }
                }
                
                $timer.Add_Tick({
                    $script:currentIndex++
                    ShowNextWord
                })
                
                # Показать первое слово сразу
                ShowNextWord
                $timer.Start()
                $displayForm.ShowDialog()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Файл пуст или не содержит слов!", "Ошибка")
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Файл не найден!", "Ошибка")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Пожалуйста, выберите файл!", "Ошибка")
    }
})
$form.Controls.Add($startButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(140, 230)
$cancelButton.Size = New-Object System.Drawing.Size(100, 30)
$cancelButton.Text = "Отмена"
$cancelButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($cancelButton)

# Показать форму
[void]$form.ShowDialog()