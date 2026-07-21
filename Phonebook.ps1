<#
.SYNOPSIS
    Телефонный справочник организации на основе Active Directory.
.DESCRIPTION
    WPF-приложение для просмотра, поиска, фильтрации,
    редактирования и экспорта контактных данных сотрудников из AD.

    Версия продукта: см. окно «О программе»
    Постановщик задачи: Арбуханов Рустам

    Разработчики:
    - Редактор: VS Code
    - AI-ассистент: MiMo Code (Xiaomi, модель mimo-auto)
    - Язык: PowerShell 5.1+ / WPF (XAML)

    Ключевые улучшения v4.0:
    - Тёмная тема (переключение одной кнопкой)
    - Автодополнение поиска (выпадающий список подсказок)
    - Ограничение редактирования по ролям (только Domain Admins)
    - 70/30 split: таблица + панель быстрого просмотра
    - Фильтр по должности (ComboBox)
    - Статус-бар с индикатором подключения к AD
    - Анимация появления окон
    - Обновлённая цветовая палитра (корпоративный стиль)
.NOTES
    Требования: PowerShell 5.1, .NET Framework 4.5+, доступ к Active Directory
    Модули: System.DirectoryServices (без модуля ActiveDirectory)
#>

#Requires -Version 5.1
Add-Type -AssemblyName PresentationFramework,PresentationCore,System.DirectoryServices,WindowsBase,System.DirectoryServices.AccountManagement

# ============================================================
# СКРЫТИЕ КОНСОЛИ POWERSHELL (чтобы не было двух окон)
# ============================================================
Add-Type -Name ConsoleTools -Namespace Win32 -MemberDefinition @'
[DllImport("Kernel32.dll", SetLastError = true)]
public static extern IntPtr GetConsoleWindow();
[DllImport("User32.dll", SetLastError = true)]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
[Win32.ConsoleTools]::ShowWindow([Win32.ConsoleTools]::GetConsoleWindow(), 0) | Out-Null

# ============================================================
# ОПРЕДЕЛЕНИЕ ВЕРСИИ (из .exe или по умолчанию)
# ============================================================
$script:appVersion = try {
    $v = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($MyInvocation.MyCommand.Path).FileVersion
    if ($v) { $v } else { "5.3" }
} catch { "5.3" }

# ============================================================
# XAML ГЛАВНОГО ОКНА (дизайн v5.0 — "Зеленое яблоко")
# 70/30 split: таблица + панель быстрого просмотра
# ============================================================
$xamlMainWindow = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Телефонный справочник" Width="1100" Height="680"
        WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="13"
        Background="{DynamicResource WindowBgBrush}"
        Opacity="0">
    <Window.Resources>
        <!-- Цвета корпоративной палитры "Зеленое яблоко" v5.0 -->
        <Color x:Key="AccentColor">#FF5EB048</Color>
        <Color x:Key="AccentHoverColor">#FF4A8F3A</Color>
        <Color x:Key="SecondaryColor">#FFD63384</Color>
        <Color x:Key="SecondaryHoverColor">#FFB82D6E</Color>
        <SolidColorBrush x:Key="SecondaryHoverBrush" Color="#FFB82D6E"/>
        <Color x:Key="ErrorColor">#FFF44336</Color>
        <Color x:Key="SuccessColor">#FF4CAF50</Color>
        <SolidColorBrush x:Key="AccentBrush" Color="#FF5EB048"/>
        <SolidColorBrush x:Key="AccentHoverBrush" Color="#FF4A8F3A"/>
        <SolidColorBrush x:Key="SecondaryBrush" Color="#FFD63384"/>
        <SolidColorBrush x:Key="ErrorBrush" Color="#FFF44336"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#FF4CAF50"/>

        <!-- Темные кисти (светлая тема по умолчанию) -->
        <SolidColorBrush x:Key="WindowBgBrush" Color="#F3F6F9"/>
        <SolidColorBrush x:Key="CardBgBrush" Color="White"/>
        <SolidColorBrush x:Key="AltRowBgBrush" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="ColHeaderBgBrush" Color="#E9ECEF"/>
        <SolidColorBrush x:Key="PlaceholderBgBrush" Color="#FAFBFC"/>
        <SolidColorBrush x:Key="RowBgBrush" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="InputBgBrush" Color="White"/>
        <SolidColorBrush x:Key="ReadOnlyBgBrush" Color="#F5F5F5"/>
        <SolidColorBrush x:Key="HoverBgBrush" Color="#E8ECF0"/>
        <SolidColorBrush x:Key="PressedBgBrush" Color="#D5D9DE"/>
        <SolidColorBrush x:Key="SelectedRowBgBrush" Color="#E8F5D9"/>
        <SolidColorBrush x:Key="CellHoverBgBrush" Color="#E8F5D9"/>
        <SolidColorBrush x:Key="RowHoverBgBrush" Color="#F1F8E9"/>
        <SolidColorBrush x:Key="ChipHoverBgBrush" Color="#E8F5D9"/>
        <SolidColorBrush x:Key="ChipBgBrush" Color="#E8F5D9"/>
        <SolidColorBrush x:Key="InputBorderBrush" Color="#D0D5DD"/>
        <SolidColorBrush x:Key="DividerBrush" Color="#E0E0E0"/>
        <SolidColorBrush x:Key="CardBorderBrush" Color="#ECF0F1"/>
        <SolidColorBrush x:Key="ReadOnlyBorderBrush" Color="#E8E8E8"/>
        <SolidColorBrush x:Key="TextPrimaryBrush" Color="#2C3E50"/>
        <SolidColorBrush x:Key="TextHeaderBrush" Color="#1A1A2E"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#7F8C8D"/>
        <SolidColorBrush x:Key="TextLabelBrush" Color="#8E99A4"/>
        <SolidColorBrush x:Key="TextMutedBrush" Color="#95A5A6"/>
        <SolidColorBrush x:Key="TextPlaceholderBrush" Color="#AAAAAA"/>
        <SolidColorBrush x:Key="TextPlaceholderSubBrush" Color="#CCCCCC"/>
        <SolidColorBrush x:Key="TextAboutBrush" Color="#555555"/>
        <SolidColorBrush x:Key="TextIconBrush" Color="#999999"/>

        <!-- Тени -->
        <DropShadowEffect x:Key="PanelShadow" BlurRadius="12" ShadowDepth="2"
                          Direction="270" Opacity="0.08" Color="#000000"/>
        <DropShadowEffect x:Key="CardShadow" BlurRadius="16" ShadowDepth="3"
                          Direction="270" Opacity="0.1" Color="#000000"/>
        <DropShadowEffect x:Key="RowShadow" BlurRadius="6" ShadowDepth="1"
                          Direction="270" Opacity="0.05" Color="#000000"/>

        <!-- ===== СТИЛИ КНОПОК ===== -->
        <!-- Базовая плоская кнопка -->
        <Style TargetType="Button" x:Key="FlatButton">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Margin" Value="2,0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}"
                                BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource HoverBgBrush}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource PressedBgBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Акцентная синяя кнопка -->
        <Style TargetType="Button" x:Key="AccentButton" BasedOn="{StaticResource FlatButton}">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource AccentHoverBrush}"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#FF3D7A2F"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Оранжевая кнопка (Редактировать) -->
        <Style TargetType="Button" x:Key="SecondaryButton" BasedOn="{StaticResource FlatButton}">
            <Setter Property="Background" Value="{StaticResource SecondaryBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Foreground" Value="White"/>
                    <Setter Property="Background" Value="{StaticResource SecondaryHoverBrush}"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#FFA0266E"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Круглая кнопка "i" (О программе) -->
        <Style TargetType="Button" x:Key="InfoButton">
            <Setter Property="Width" Value="36"/>
            <Setter Property="Height" Value="36"/>
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="18"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                CornerRadius="18"
                                Width="36" Height="36"
                                BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#5EB048"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#4A8F3A"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ===== СТИЛИ ТЕКСТА ===== -->
        <Style TargetType="TextBlock" x:Key="HeaderText">
            <Setter Property="FontSize" Value="18"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="{DynamicResource TextHeaderBrush}"/>
        </Style>

        <Style TargetType="TextBlock" x:Key="SubHeaderText">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="Light"/>
            <Setter Property="Foreground" Value="{DynamicResource TextSecondaryBrush}"/>
        </Style>

        <Style TargetType="TextBlock" x:Key="LabelText">
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="{DynamicResource TextLabelBrush}"/>
            <Setter Property="TextBlock.LineHeight" Value="16"/>
        </Style>

        <!-- Стиль для переноса текста в ячейках таблицы -->
        <Style x:Key="WrapText" TargetType="TextBlock">
            <Setter Property="TextWrapping" Value="Wrap"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>

        <!-- Стиль для заголовков колонок (UPPERCASE + letter-spacing) -->
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{DynamicResource ColHeaderBgBrush}"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="Height" Value="38"/>
            <Setter Property="Padding" Value="14,4"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
        </Style>

        <!-- Стиль ячеек -->
        <Style TargetType="DataGridCell">
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="TextBlock.TextWrapping" Value="Wrap"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border x:Name="cellBorder"
                                Background="Transparent"
                                Padding="{TemplateBinding Padding}"
                                BorderThickness="0"
                                SnapsToDevicePixels="True">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="cellBorder" Property="Background" Value="{DynamicResource CellHoverBgBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Стиль строки таблицы -->
        <Style TargetType="DataGridRow" x:Key="DataGridRowGreen">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{DynamicResource SelectedRowBgBrush}"/>
                    <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource RowHoverBgBrush}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <!-- Анимация появления окна -->
    <Window.Triggers>
        <EventTrigger RoutedEvent="Loaded">
            <BeginStoryboard>
                <Storyboard>
                    <DoubleAnimation Storyboard.TargetProperty="Opacity"
                                     From="0" To="1" Duration="0:0:0.3"/>
                </Storyboard>
            </BeginStoryboard>
        </EventTrigger>
    </Window.Triggers>

    <!-- ================================================================ -->
    <!-- ОСНОВНАЯ СЕТКА -->
    <!-- ================================================================ -->
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>   <!-- Шапка -->
            <RowDefinition Height="Auto"/>   <!-- Поиск + фильтры + кнопки -->
            <RowDefinition Height="*"/>      <!-- Таблица + панель просмотра -->
            <RowDefinition Height="Auto"/>   <!-- Статус-бар -->
        </Grid.RowDefinitions>

        <!-- ============================================================ -->
        <!-- СТРОКА 0: ШАПКА -->
        <!-- ============================================================ -->
        <Border Grid.Row="0" Margin="0,0,0,8" Padding="16,12"
                Background="{DynamicResource CardBgBrush}" CornerRadius="10"
                Effect="{StaticResource PanelShadow}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Логотип (изображение) -->
                <Image Grid.Column="0" Name="LogoImage"
                       Width="44" Height="44"
                       Margin="0,0,12,0"
                       RenderOptions.BitmapScalingMode="HighQuality"
                       ToolTip="Зеленое яблоко"/>

                <!-- Заголовки -->
                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock Style="{StaticResource HeaderText}">
                        <Run Text="Зеленое яблоко"/><Run Text=" — "/><Run Text="Телефонный справочник" FontWeight="Normal" FontSize="14" Foreground="{DynamicResource TextSecondaryBrush}"/>
                    </TextBlock>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Name="WebsiteLink" FontSize="11" Cursor="Hand"
                                   Foreground="{DynamicResource AccentBrush}"
                                   ToolTip="Открыть сайт zelenoeyabloko.ru">
                            <Run Text="🌐"/><Run Text="  "/><Run Text="zelenoeyabloko.ru"/>
                        </TextBlock>
                        <TextBlock Text="  |  " FontSize="11" Foreground="{DynamicResource TextMutedBrush}"/>
                        <TextBlock Name="FeedbackEmail" FontSize="11" Cursor="Hand"
                                   Foreground="{DynamicResource AccentBrush}"
                                   ToolTip="Написать на почту поддержки">
                            <Run Text="✉️"/><Run Text="  "/><Run Text="it@pepper-group.ru"/>
                        </TextBlock>
                    </StackPanel>
                </StackPanel>

                <!-- Кнопки управления -->
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Name="ThemeToggleButton" Content="🌙"
                            Style="{StaticResource InfoButton}"
                            Margin="0,0,8,0"
                            ToolTip="Переключить тему"/>
                    <Button Name="AboutButton" Content="i"
                            Style="{StaticResource InfoButton}"
                            ToolTip="О программе"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ============================================================ -->
        <!-- СТРОКА 1: ПОИСК + ФИЛЬТРЫ + КНОПКИ -->
        <!-- ============================================================ -->
        <Border Grid.Row="1" Margin="0,0,0,8" Padding="12,8"
                Background="{DynamicResource PlaceholderBgBrush}" CornerRadius="10"
                Effect="{StaticResource PanelShadow}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="350"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Поле поиска с плейсхолдером -->
                <Border Grid.Column="0" CornerRadius="8" BorderThickness="1"
                        BorderBrush="{DynamicResource InputBorderBrush}" Background="{DynamicResource InputBgBrush}"
                        Height="34">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="🔍" FontSize="11"
                                   VerticalAlignment="Center" Margin="10,0,6,0" Foreground="{DynamicResource TextIconBrush}"/>
                        <TextBox Name="SearchBox" Grid.Column="1"
                                 BorderThickness="0" Padding="2,0"
                                 Background="Transparent"
                                 FontSize="13" Foreground="{DynamicResource TextPrimaryBrush}"
                                 VerticalAlignment="Center"
                                 ToolTip="Поиск по фамилии, имени, отделу, должности, телефону, email"/>
                        <!-- Плейсхолдер -->
                        <TextBlock Name="SearchPlaceholder" Grid.Column="1"
                                   Text="Поиск по имени, фамилии..."
                                   FontSize="13" Foreground="{DynamicResource TextPlaceholderBrush}"
                                   VerticalAlignment="Center"
                                   IsHitTestVisible="False"
                                   Margin="2,0,0,0"/>
                        <Button Name="ClearButton" Grid.Column="2"
                                Content="✕" Width="24" Height="24"
                                Background="Transparent" Foreground="{DynamicResource TextIconBrush}"
                                BorderThickness="0" Cursor="Hand"
                                Margin="0,0,4,0" Padding="0"
                                FontSize="10"
                                Visibility="Collapsed"/>
                    </Grid>
                </Border>

                <!-- Выпадающий список подсказок поиска -->
                <Popup Name="SearchSuggestions" PlacementTarget="{Binding ElementName=SearchBox}"
                       Placement="Bottom" StaysOpen="False" AllowsTransparency="True"
                       PopupAnimation="Fade" MaxHeight="200">
                    <Border Background="White" BorderBrush="#D0D5DD" BorderThickness="1"
                            CornerRadius="8" Margin="0,4,0,0"
                            Effect="{StaticResource CardShadow}" MinWidth="316">
                        <ListBox Name="SuggestionsList" BorderThickness="0"
                                 Background="Transparent" MaxHeight="180"
                                 ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                                 FontSize="12">
                            <ListBox.ItemContainerStyle>
                                <Style TargetType="ListBoxItem">
                                    <Setter Property="Padding" Value="10,6"/>
                                    <Setter Property="Cursor" Value="Hand"/>
                                    <Setter Property="Template">
                                        <Setter.Value>
                                            <ControlTemplate TargetType="ListBoxItem">
                                                <Border x:Name="b" Padding="{TemplateBinding Padding}"
                                                        Background="Transparent" CornerRadius="4" Margin="2">
                                                    <ContentPresenter/>
                                                </Border>
                                                <ControlTemplate.Triggers>
                                                    <Trigger Property="IsMouseOver" Value="True">
                                                        <Setter TargetName="b" Property="Background" Value="#E8F5D9"/>
                                                    </Trigger>
                                                </ControlTemplate.Triggers>
                                            </ControlTemplate>
                                        </Setter.Value>
                                    </Setter>
                                </Style>
                            </ListBox.ItemContainerStyle>
                        </ListBox>
                    </Border>
                </Popup>

                <!-- Фильтр по отделу -->
                <StackPanel Grid.Column="1" Orientation="Horizontal"
                            VerticalAlignment="Center" Margin="8,0,8,0">
                    <TextBlock Text="🏢" FontSize="11" VerticalAlignment="Center" Margin="0,0,4,0"/>
                    <ComboBox Name="DepartmentFilter" Width="200" Height="32"
                              FontSize="12"/>
                </StackPanel>

                <!-- Фильтр по должности -->
                <StackPanel Grid.Column="2" Orientation="Horizontal"
                            VerticalAlignment="Center" Margin="0,0,12,0">
                    <TextBlock Text="💼" FontSize="11" VerticalAlignment="Center" Margin="0,0,4,0"/>
                    <ComboBox Name="TitleFilter" Width="200" Height="32"
                              FontSize="12"/>
                </StackPanel>

                <!-- Кнопка "Создать новый контакт" -->
                <Button Name="CreateButton" Grid.Column="3"
                        Content="＋ Создать"
                        Style="{StaticResource AccentButton}"
                        Height="32" Margin="0,0,8,0"
                        ToolTip="Создать нового сотрудника"/>

                <!-- Кнопка "Экспорт CSV" -->
                <Button Name="ExportButton" Grid.Column="3"
                        Content="Экспорт"
                        Style="{StaticResource AccentButton}"
                        Height="32" Margin="0,0,8,0"
                        ToolTip="Сохранить отфильтрованные данные в CSV"/>

                <!-- Счётчик-чип -->
                <Border Grid.Column="4" CornerRadius="14" Padding="14,5,14,5"
                        Background="{StaticResource AccentBrush}"
                        Margin="8,0,0,0">
                    <TextBlock Name="CountText" FontSize="11" FontWeight="SemiBold"
                               Foreground="White" VerticalAlignment="Center"
                               LineHeight="16"/>
                </Border>
            </Grid>
        </Border>

        <!-- ============================================================ -->
        <!-- СТРОКА 2: ТАБЛИЦА (70%) + ПАНЕЛЬ ПРОСМОТРА (30%) -->
        <!-- ============================================================ -->
        <Grid Grid.Row="2" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="7*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="3*"/>
            </Grid.ColumnDefinitions>

            <!-- ТАБЛИЦА (левая часть, 70%) -->
            <Border Grid.Column="0" CornerRadius="10" Padding="0"
                    Background="{DynamicResource CardBgBrush}" Effect="{StaticResource PanelShadow}">
                <DataGrid Name="DataGrid"
                          AutoGenerateColumns="False" IsReadOnly="True"
                          SelectionMode="Single" SelectionUnit="FullRow"
                          CanUserSortColumns="True" CanUserResizeColumns="True"
                          CanUserReorderColumns="True"
                          GridLinesVisibility="None"
                          BorderThickness="0"
                          FontSize="13"
                          RowHeight="60"
                          FontFamily="Segoe UI"
                          AlternatingRowBackground="{DynamicResource AltRowBgBrush}"
                          RowBackground="{DynamicResource RowBgBrush}"
                          HeadersVisibility="Column"
                          Margin="0"
                          VirtualizingPanel.VirtualizationMode="Recycling"
                          VirtualizingPanel.ScrollUnit="Pixel">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="ФАМИЛИЯ"   Binding="{Binding Surname}"        Width="120"/>
                        <DataGridTextColumn Header="ИМЯ"       Binding="{Binding GivenName}"       Width="100"/>
                        <DataGridTextColumn Header="ОТДЕЛ"     Binding="{Binding Department}"      Width="*" MinWidth="150" ElementStyle="{StaticResource WrapText}"/>
                        <DataGridTextColumn Header="ДОЛЖНОСТЬ" Binding="{Binding Title}"           Width="*" MinWidth="150" ElementStyle="{StaticResource WrapText}"/>
                        <DataGridTextColumn Header="ТЕЛ."      Binding="{Binding TelephoneNumber}" Width="70"/>
                        <DataGridTextColumn Header="EMAIL"     Binding="{Binding Mail}"            Width="160" ElementStyle="{StaticResource WrapText}"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Border>

            <!-- Разделитель -->
            <GridSplitter Grid.Column="1" Width="4"
                          HorizontalAlignment="Center"
                          VerticalAlignment="Stretch"
                          Background="Transparent"
                          Cursor="SizeWE"
                          Margin="4,4,4,4"/>

            <!-- ПАНЕЛЬ БЫСТРОГО ПРОСМОТРА (правая часть, 30%) -->
            <Border Name="PreviewPanel" Grid.Column="2" CornerRadius="10"
                    Background="{DynamicResource CardBgBrush}" Effect="{StaticResource PanelShadow}"
                    Padding="14">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>  <!-- Аватар -->
                        <RowDefinition Height="Auto"/>  <!-- ФИО -->
                        <RowDefinition Height="Auto"/>  <!-- Должность -->
                        <RowDefinition Height="Auto"/>  <!-- Отдел -->
                        <RowDefinition Height="Auto"/>  <!-- Разделитель -->
                        <RowDefinition Height="Auto"/>  <!-- Телефон -->
                        <RowDefinition Height="Auto"/>  <!-- Email -->
                        <RowDefinition Height="*"/>     <!-- Заполнитель -->
                        <RowDefinition Height="Auto"/>  <!-- Кнопки -->
                    </Grid.RowDefinitions>

                    <!-- Заглушка когда ничего не выбрано -->
                    <Border Name="PreviewPlaceholder" Grid.RowSpan="9"
                            Background="{DynamicResource PlaceholderBgBrush}"
                            CornerRadius="8"
                            VerticalAlignment="Center">
                        <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center"
                                    Margin="20">
                            <TextBlock Text="👤" FontSize="52"
                                       HorizontalAlignment="Center" Opacity="0.25"/>
                            <TextBlock Text="Выберите сотрудника"
                                       FontSize="14" FontWeight="Medium"
                                       Foreground="{DynamicResource TextSecondaryBrush}"
                                       HorizontalAlignment="Center" Margin="0,10,0,0"/>
                            <TextBlock Text="для просмотра контактов"
                                       FontSize="12" Foreground="{DynamicResource TextMutedBrush}"
                                       HorizontalAlignment="Center"/>
                        </StackPanel>
                    </Border>

                    <!-- Аватар (инициалы) -->
                    <Border Name="AvatarBorder" Grid.Row="0"
                            Width="72" Height="72" CornerRadius="36"
                            Background="{StaticResource AccentBrush}"
                            HorizontalAlignment="Center"
                            Margin="0,8,0,16"
                            Visibility="Collapsed">
                        <TextBlock Name="AvatarText" FontSize="24" FontWeight="Bold"
                                   Foreground="White"
                                   HorizontalAlignment="Center"
                                   VerticalAlignment="Center"/>
                    </Border>

                    <!-- ФИО -->
                    <TextBlock Name="PreviewName" Grid.Row="1"
                               FontSize="16" FontWeight="SemiBold" Foreground="{DynamicResource TextHeaderBrush}"
                               HorizontalAlignment="Center"
                               TextAlignment="Center"
                               TextWrapping="Wrap"
                               Visibility="Collapsed"/>

                    <!-- Должность -->
                    <StackPanel Grid.Row="2" Orientation="Horizontal"
                                HorizontalAlignment="Center"
                                Margin="0,4,0,0"
                                Visibility="Collapsed" Name="PreviewTitleRow">
                        <TextBlock Name="PreviewTitle" FontSize="12" Foreground="{DynamicResource TextSecondaryBrush}"
                                   HorizontalAlignment="Center"/>
                    </StackPanel>

                    <!-- Отдел -->
                    <StackPanel Grid.Row="3" Orientation="Horizontal"
                                HorizontalAlignment="Center"
                                Margin="0,2,0,0"
                                Visibility="Collapsed" Name="PreviewDeptRow">
                        <TextBlock Text="🏢 " FontSize="11"/>
                        <TextBlock Name="PreviewDept" FontSize="11" Foreground="{DynamicResource TextMutedBrush}"/>
                    </StackPanel>

                    <!-- Разделитель -->
                    <Rectangle Grid.Row="4" Height="1" Fill="{DynamicResource CardBorderBrush}"
                               Margin="0,12,0,12"
                               Visibility="Collapsed" Name="PreviewDivider"/>

                    <!-- Телефон -->
                    <Border Grid.Row="5" CornerRadius="8" Padding="12,8"
                            Background="{DynamicResource RowBgBrush}" BorderThickness="1" BorderBrush="{DynamicResource CardBorderBrush}"
                            Margin="0,0,0,8"
                            Visibility="Collapsed" Name="PreviewPhoneRow">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" Text="📞 " FontSize="13"
                                       VerticalAlignment="Center"/>
                            <StackPanel Grid.Column="1" Margin="6,0,0,0"
                                        VerticalAlignment="Center">
                                <TextBlock Text="Телефон" Style="{StaticResource LabelText}"/>
                                <TextBlock Name="PreviewPhone" FontSize="13" FontWeight="SemiBold"
                                           Foreground="{DynamicResource TextPrimaryBrush}"/>
                            </StackPanel>
                            <Button Name="PreviewCopyPhone" Grid.Column="2"
                                    Content="📋" Width="30" Height="30"
                                    Background="Transparent" BorderThickness="0"
                                    Cursor="Hand" FontSize="14"
                                    ToolTip="Копировать телефон"/>
                        </Grid>
                    </Border>

                    <!-- Email -->
                    <Border Grid.Row="6" CornerRadius="8" Padding="12,8"
                            Background="{DynamicResource RowBgBrush}" BorderThickness="1" BorderBrush="{DynamicResource CardBorderBrush}"
                            Margin="0,0,0,8"
                            Visibility="Collapsed" Name="PreviewEmailRow">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" Text="✉️ " FontSize="13"
                                       VerticalAlignment="Center"/>
                            <StackPanel Grid.Column="1" Margin="6,0,0,0"
                                        VerticalAlignment="Center">
                                <TextBlock Text="Email" Style="{StaticResource LabelText}"/>
                                <TextBlock Name="PreviewEmail" FontSize="12"
                                           Foreground="{DynamicResource TextPrimaryBrush}"
                                           TextTrimming="CharacterEllipsis"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- Кнопка "Редактировать" -->
                    <Button Name="PreviewEditButton" Grid.Row="8"
                            Content="✏️  Редактировать"
                            Style="{StaticResource AccentButton}"
                            HorizontalAlignment="Stretch"
                            Height="36" Margin="0,4,0,0"
                            Visibility="Collapsed"
                            ToolTip="Редактировать данные сотрудника"/>
                </Grid>
            </Border>
        </Grid>

        <!-- ============================================================ -->
        <!-- СТРОКА 3: СТАТУС-БАР -->
        <!-- ============================================================ -->
        <Border Grid.Row="3" Margin="0,0,0,6" Padding="12,6"
                Background="{DynamicResource CardBgBrush}" CornerRadius="8"
                Effect="{StaticResource PanelShadow}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Индикатор подключения к AD -->
                <Ellipse Name="StatusIndicator" Grid.Column="0"
                         Width="8" Height="8" Fill="#BDC3C7" Margin="4,0,6,0"
                         VerticalAlignment="Center"/>

                <!-- Статус подключения -->
                <TextBlock Name="StatusText" Grid.Column="1"
                           VerticalAlignment="Center"
                           FontSize="11" Foreground="{DynamicResource TextLabelBrush}"/>

                <!-- Давность данных -->
                <TextBlock Grid.Column="3" FontSize="11" Foreground="{DynamicResource TextLabelBrush}"
                           VerticalAlignment="Center"
                           Name="LastUpdateText"/>

                <!-- Вертикальный разделитель -->
                <Rectangle Grid.Column="4" Width="1" Height="14"
                           Fill="{DynamicResource DividerBrush}" Margin="8,0"/>

                <!-- Счётчик записей -->
                <TextBlock Name="TotalCountText" Grid.Column="5"
                           FontSize="11" Foreground="{DynamicResource TextLabelBrush}"
                           VerticalAlignment="Center"/>

                <!-- Вертикальный разделитель -->
                <Rectangle Grid.Column="6" Width="1" Height="14"
                           Fill="{DynamicResource DividerBrush}" Margin="8,0"/>

                <!-- Пользователь -->
                <TextBlock Name="CurrentUserText" Grid.Column="7"
                           FontSize="11" Foreground="{DynamicResource TextLabelBrush}"
                           VerticalAlignment="Center"/>
            </Grid>
        </Border>

        <!-- ============================================================ -->
        <!-- СТРОКА 4: УДАЛЕНА (кнопки перенесены в панель фильтров) -->
        <!-- ============================================================ -->
    </Grid>
    <!-- КОНЕЦ ОСНОВНОЙ СЕТКИ -->
</Window>
'@

# ============================================================
# XAML КАРТОЧКИ СОТРУДНИКА (две колонки, современный дизайн)
# ============================================================
$xamlCardWindow = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        MinWidth="520" MinHeight="580"
        Width="540" Height="Auto" SizeToContent="Height"
        WindowStartupLocation="CenterOwner"
        ResizeMode="CanResizeWithGrip" FontFamily="Segoe UI" FontSize="13"
        Background="#F3F6F9" UseLayoutRounding="True"
        Opacity="0">
    <Window.Resources>
        <Style TargetType="TextBlock" x:Key="FieldLabel">
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#8E99A4"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="HorizontalAlignment" Value="Right"/>
        </Style>

        <!-- Стиль поля ввода -->
        <Style TargetType="TextBox" x:Key="FieldBox">
            <Setter Property="Margin" Value="0,0,0,8"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#D0D5DD"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Foreground" Value="#2C3E50"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="border" CornerRadius="6"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
                            <ScrollViewer x:Name="PART_ContentHost"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#5EB048"/>
                            </Trigger>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#5EB048"/>
                                <Setter TargetName="border" Property="BorderThickness" Value="2"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Стиль поля только для чтения (серый фон, read-only) -->
        <Style TargetType="TextBox" x:Key="ReadOnlyBox" BasedOn="{StaticResource FieldBox}">
            <Setter Property="Background" Value="#F5F5F5"/>
            <Setter Property="BorderBrush" Value="#E8E8E8"/>
            <Setter Property="Foreground" Value="#7F8C8D"/>
            <Setter Property="IsReadOnly" Value="True"/>
        </Style>

        <!-- Стиль кнопок карточки -->
        <Style TargetType="Button" x:Key="CardButton">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="Margin" Value="4,0"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}"
                                BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.85"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.7"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <DropShadowEffect x:Key="CardShadow" BlurRadius="16" ShadowDepth="3"
                          Direction="270" Opacity="0.12" Color="#000000"/>
    </Window.Resources>

    <Window.Triggers>
        <EventTrigger RoutedEvent="Loaded">
            <BeginStoryboard>
                <Storyboard>
                    <DoubleAnimation Storyboard.TargetProperty="Opacity"
                                     From="0" To="1" Duration="0:0:0.25"/>
                </Storyboard>
            </BeginStoryboard>
        </EventTrigger>
    </Window.Triggers>

    <Border Margin="14" CornerRadius="12"
            Background="#F3F6F9"
            Effect="{StaticResource CardShadow}">
        <Grid Margin="22,18">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>   <!-- 0: Заголовок -->
                <RowDefinition Height="Auto"/>   <!-- 1: Имя -->
                <RowDefinition Height="Auto"/>   <!-- 2: Фамилия -->
                <RowDefinition Height="Auto"/>   <!-- 3: Должность -->
                <RowDefinition Height="Auto"/>   <!-- 4: Отдел -->
                <RowDefinition Height="Auto"/>   <!-- 5: Телефон -->
                <RowDefinition Height="Auto"/>   <!-- 6: Email -->
                <RowDefinition Height="Auto"/>   <!-- 7: Логин -->
                <RowDefinition Height="*"/>      <!-- 8: Заполнитель (чтобы кнопки были внизу) -->
                <RowDefinition Height="Auto"/>   <!-- 9: Панель кнопок -->
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="110"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Заголовок с иконкой и текстом (StackPanel — вертикальное расположение, текст переносится) -->
            <StackPanel Grid.Row="0" Grid.ColumnSpan="2"
                       Orientation="Vertical" HorizontalAlignment="Center" Margin="0,0,0,18">
                <Border Width="52" Height="52" CornerRadius="26"
                        Background="#EAF2FA" HorizontalAlignment="Center">
                    <TextBlock Text="👤" FontSize="24"
                              HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <TextBlock Name="CardTitle"
                          FontSize="16" FontWeight="Bold" Foreground="#1A1A2E"
                          HorizontalAlignment="Center"
                          TextWrapping="Wrap" TextTrimming="CharacterEllipsis"
                          MaxWidth="420"
                          Margin="0,10,0,0"/>
            </StackPanel>

            <!-- Поле: Имя (только чтение) -->
            <TextBlock Grid.Row="1" Grid.Column="0" Text="👤  Имя:" Style="{StaticResource FieldLabel}"/>
            <TextBox Grid.Row="1" Grid.Column="1" Name="GivenNameBox" Style="{StaticResource ReadOnlyBox}"/>

            <!-- Поле: Фамилия (только чтение) -->
            <TextBlock Grid.Row="2" Grid.Column="0" Text="👤  Фамилия:" Style="{StaticResource FieldLabel}"/>
            <TextBox Grid.Row="2" Grid.Column="1" Name="SurnameBox" Style="{StaticResource ReadOnlyBox}"/>

            <!-- Поле: Должность -->
            <TextBlock Grid.Row="3" Grid.Column="0" Text="💼  Должность:" Style="{StaticResource FieldLabel}"/>
            <TextBox Grid.Row="3" Grid.Column="1" Name="TitleBox" Style="{StaticResource FieldBox}"/>

            <!-- Поле: Отдел -->
            <TextBlock Grid.Row="4" Grid.Column="0" Text="🏢  Отдел:" Style="{StaticResource FieldLabel}"/>
            <TextBox Grid.Row="4" Grid.Column="1" Name="DepartmentBox" Style="{StaticResource FieldBox}"/>

            <!-- Поле: Телефон -->
            <TextBlock Grid.Row="5" Grid.Column="0" Text="📞  Телефон:" Style="{StaticResource FieldLabel}"/>
            <TextBox Grid.Row="5" Grid.Column="1" Name="TelephoneBox" Style="{StaticResource FieldBox}"/>

            <!-- Поле: Email -->
            <TextBlock Grid.Row="6" Grid.Column="0" Text="✉️  Email:" Style="{StaticResource FieldLabel}"/>
            <TextBox Grid.Row="6" Grid.Column="1" Name="MailBox" Style="{StaticResource FieldBox}"/>

            <!-- Поле: Логин (только чтение) -->
            <TextBlock Grid.Row="7" Grid.Column="0" Text="🔑  Логин:" Style="{StaticResource FieldLabel}"/>
            <TextBox Grid.Row="7" Grid.Column="1" Name="LoginBox" Style="{StaticResource ReadOnlyBox}"/>

            <!-- ============================================================ -->
            <!-- ПАНЕЛЬ КНОПОК (всегда внизу, строка 9) -->
            <!-- ============================================================ -->
            <Border Grid.Row="9" Grid.ColumnSpan="2"
                    Margin="0,16,0,0" Padding="0,12,0,0"
                    BorderBrush="#E0E0E0" BorderThickness="0,1,0,0">
                <WrapPanel HorizontalAlignment="Right">
                    <Button Name="CopyPhoneButton" Content="📋  Копировать телефон"
                            Style="{StaticResource CardButton}"
                            Background="#5EB048" Foreground="White"/>
                    <Button Name="SaveButton" Content="💾  Сохранить"
                            Style="{StaticResource CardButton}"
                            Background="#4CAF50" Foreground="White"
                            Visibility="Collapsed"/>
                    <Button Name="CloseButton" Content="✕  Закрыть"
                            Style="{StaticResource CardButton}"
                            Background="#E0E0E0" Foreground="#2C3E50"/>
                </WrapPanel>
            </Border>
        </Grid>
    </Border>
</Window>
'@

# ============================================================
# XAML ДИАЛОГА ВХОДА (обязательная аутентификация при запуске)
# ============================================================
$xamlLoginDialog = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="420" Height="Auto" SizeToContent="Height"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize" FontFamily="Segoe UI" FontSize="13"
        Background="#F3F6F9" Title="Вход — Телефонный справочник" Opacity="0">
    <Window.Resources>
        <DropShadowEffect x:Key="CardShadow" BlurRadius="16" ShadowDepth="3"
                          Direction="270" Opacity="0.12" Color="#000000"/>
    </Window.Resources>
    <Window.Triggers>
        <EventTrigger RoutedEvent="Loaded">
            <BeginStoryboard>
                <Storyboard>
                    <DoubleAnimation Storyboard.TargetProperty="Opacity"
                                     From="0" To="1" Duration="0:0:0.2"/>
                </Storyboard>
            </BeginStoryboard>
        </EventTrigger>
    </Window.Triggers>
    <Border Margin="14" CornerRadius="12" Background="#F3F6F9"
            Effect="{StaticResource CardShadow}">
        <Border Margin="20,16" CornerRadius="8" Background="White" Padding="20,16">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Row="0" Grid.ColumnSpan="2" Margin="0,0,0,14">
                    <TextBlock Text="&#x1F512;  Вход в справочник"
                               FontSize="14" FontWeight="Bold" Foreground="#2C3E50"/>
                    <TextBlock Text="Требуются учётные данные доменного администратора"
                               FontSize="11" Foreground="#7F8C8D" TextWrapping="Wrap" Margin="0,4,0,0"/>
                </StackPanel>
                <TextBlock Grid.Row="1" Grid.Column="0" Text="&#x1F464;  Логин:" Margin="0,0,0,8"
                           FontSize="11" FontWeight="SemiBold" Foreground="#8E99A4" VerticalAlignment="Center"/>
                <TextBox Grid.Row="1" Grid.Column="1" Name="LoginBox"
                         Margin="0,0,0,8" Padding="10,6" BorderThickness="1"
                         BorderBrush="#D0D5DD" Background="White" FontSize="13" Foreground="#2C3E50"/>
                <TextBlock Grid.Row="2" Grid.Column="0" Text="&#x1F511;  Пароль:" Margin="0,0,0,8"
                           FontSize="11" FontWeight="SemiBold" Foreground="#8E99A4" VerticalAlignment="Center"/>
                <PasswordBox Grid.Row="2" Grid.Column="1" Name="LoginPasswordBox"
                             Margin="0,0,0,8" Padding="10,6" BorderThickness="1"
                             BorderBrush="#D0D5DD" Background="White" FontSize="13"/>
                <TextBlock Grid.Row="3" Grid.ColumnSpan="2" Name="LoginErrorText"
                           Foreground="#F44336" FontSize="11" TextWrapping="Wrap"
                           Margin="0,0,0,8" Visibility="Collapsed"/>
                <TextBlock Grid.Row="4" Grid.ColumnSpan="2" Name="LoginStatusText"
                           Foreground="#757575" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,12"/>
                <WrapPanel Grid.Row="5" Grid.ColumnSpan="2" HorizontalAlignment="Right">
                    <Button Name="LoginOkButton" Content="&#x2714;  Войти"
                            Background="#5EB048" Foreground="White"
                            FontSize="12" FontWeight="SemiBold" Padding="16,7"
                            BorderThickness="0" Cursor="Hand" Margin="0,0,6,0"/>
                    <Button Name="LoginCancelButton" Content="&#x2716;  Отмена"
                            Background="#E0E0E0" Foreground="#2C3E50"
                            FontSize="12" FontWeight="SemiBold" Padding="16,7"
                            BorderThickness="0" Cursor="Hand"/>
                </WrapPanel>
            </Grid>
        </Border>
    </Border>
</Window>
'@

# ============================================================
# XAML ДИАЛОГА ПАРОЛЯ АДМИНИСТРАТОРА
# ============================================================
$xamlPasswordDialog = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="400" Height="Auto" SizeToContent="Height"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize" FontFamily="Segoe UI" FontSize="13"
        Background="#F3F6F9" Title="Подтверждение прав" Opacity="0">
    <Window.Resources>
        <DropShadowEffect x:Key="CardShadow" BlurRadius="16" ShadowDepth="3"
                          Direction="270" Opacity="0.12" Color="#000000"/>
    </Window.Resources>
    <Window.Triggers>
        <EventTrigger RoutedEvent="Loaded">
            <BeginStoryboard>
                <Storyboard>
                    <DoubleAnimation Storyboard.TargetProperty="Opacity"
                                     From="0" To="1" Duration="0:0:0.2"/>
                </Storyboard>
            </BeginStoryboard>
        </EventTrigger>
    </Window.Triggers>
    <Border Margin="14" CornerRadius="12" Background="#F3F6F9"
            Effect="{StaticResource CardShadow}">
        <Border Margin="20,16" CornerRadius="8" Background="White" Padding="20,16">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <!-- Header -->
                <StackPanel Grid.Row="0" Grid.ColumnSpan="2" Margin="0,0,0,14">
                    <TextBlock Text="&#x1F512;  Ввод пароля администратора"
                               FontSize="14" FontWeight="Bold" Foreground="#2C3E50"/>
                    <TextBlock Text="Для редактирования данныхrequires доменный администратор"
                               FontSize="11" Foreground="#7F8C8D" TextWrapping="Wrap" Margin="0,4,0,0"/>
                </StackPanel>

                <!-- Login -->
                <TextBlock Grid.Row="1" Grid.Column="0" Text="&#x1F464;  Логин:" Margin="0,0,0,8"
                           FontSize="11" FontWeight="SemiBold" Foreground="#8E99A4"
                           VerticalAlignment="Center"/>
                <TextBox Grid.Row="1" Grid.Column="1" Name="AdminLoginBox"
                         Margin="0,0,0,8" Padding="10,6" BorderThickness="1"
                         BorderBrush="#D0D5DD" Background="White" FontSize="13"
                         Foreground="#2C3E50"/>

                <!-- Password -->
                <TextBlock Grid.Row="2" Grid.Column="0" Text="&#x1F511;  Пароль:" Margin="0,0,0,8"
                           FontSize="11" FontWeight="SemiBold" Foreground="#8E99A4"
                           VerticalAlignment="Center"/>
                <PasswordBox Grid.Row="2" Grid.Column="1" Name="AdminPasswordBox"
                             Margin="0,0,0,8" Padding="10,6" BorderThickness="1"
                             BorderBrush="#D0D5DD" Background="White" FontSize="13"/>

                <!-- Error text -->
                <TextBlock Grid.Row="3" Grid.ColumnSpan="2" Name="AuthErrorText"
                           Foreground="#F44336" FontSize="11" TextWrapping="Wrap"
                           Margin="0,0,0,8" Visibility="Collapsed"/>

                <!-- Status -->
                <TextBlock Grid.Row="4" Grid.ColumnSpan="2" Name="AuthStatusText"
                           Foreground="#757575" FontSize="11" TextWrapping="Wrap"
                           Margin="0,0,0,12"/>

                <!-- Buttons -->
                <WrapPanel Grid.Row="5" Grid.ColumnSpan="2" HorizontalAlignment="Right">
                    <Button Name="AuthOkButton" Content="&#x2714;  Войти"
                            Background="#5EB048" Foreground="White"
                            FontSize="12" FontWeight="SemiBold" Padding="16,7"
                            BorderThickness="0" Cursor="Hand" Margin="0,0,6,0"/>
                    <Button Name="AuthCancelButton" Content="&#x2716;  Отмена"
                            Background="#E0E0E0" Foreground="#2C3E50"
                            FontSize="12" FontWeight="SemiBold" Padding="16,7"
                            BorderThickness="0" Cursor="Hand"/>
                </WrapPanel>
            </Grid>
        </Border>
    </Border>
</Window>
'@

# ============================================================
# XAML ОКНА «О ПРОГРАММЕ»
# ============================================================
$xamlAboutWindow = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="520" Height="480" WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize" FontFamily="Segoe UI" FontSize="13"
        Background="#F8F9FA" UseLayoutRounding="True"
        Opacity="0">
    <Window.Resources>
        <DropShadowEffect x:Key="CardShadow" BlurRadius="16" ShadowDepth="3"
                          Direction="270" Opacity="0.12" Color="#000000"/>

        <Style TargetType="Button" x:Key="AboutCloseButton">
            <Setter Property="Background" Value="#5EB048"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="24,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}"
                                BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#4A8F3A"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#3D7A2F"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Window.Triggers>
        <EventTrigger RoutedEvent="Loaded">
            <BeginStoryboard>
                <Storyboard>
                    <DoubleAnimation Storyboard.TargetProperty="Opacity"
                                     From="0" To="1" Duration="0:0:0.25"/>
                </Storyboard>
            </BeginStoryboard>
        </EventTrigger>
    </Window.Triggers>

    <Border Margin="14" CornerRadius="12"
            Background="#F8F9FA"
            Effect="{StaticResource CardShadow}">
        <Border Margin="20,16" CornerRadius="8" Background="White" Padding="20,16">
            <DockPanel>
                <!-- Кнопка Закрыть (всегда внизу) -->
                <Button DockPanel.Dock="Bottom" Name="AboutCloseButton"
                        Content="✕  Закрыть" Style="{StaticResource AboutCloseButton}"
                        HorizontalAlignment="Center" Margin="0,10,0,0"
                        ToolTip="Закрыть окно"/>

                <!-- Контент с прокруткой -->
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel>
                        <!-- Заголовок -->
                        <StackPanel Margin="0,0,0,6">
                            <TextBlock Text="📋  Зеленое яблоко — Телефонный справочник"
                                       FontSize="15" FontWeight="Bold" Foreground="#5EB048"
                                       TextWrapping="Wrap"/>
                        </StackPanel>

                        <!-- Версия -->
                        <Grid Margin="0,0,0,12">
                            <TextBlock FontSize="12" Foreground="#555">
                                <Run FontWeight="SemiBold" Text="Версия:"/>
                                <Run Text=" "/><Run Name="VersionText" Text="1.0.0"/>
                            </TextBlock>
                        </Grid>

                        <!-- Разделитель -->
                        <Rectangle Height="1" Fill="#ECF0F1" Margin="0,0,0,10"/>

                        <!-- Проделанная работа -->
                        <StackPanel Margin="0,0,0,10">
                            <TextBlock Text="✅  Проделанная работа" FontSize="13"
                                       FontWeight="SemiBold" Foreground="#2C3E50" Margin="0,0,0,6"/>
                            <BulletDecorator Margin="0,0,0,3">
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="14" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="Создано приложение для поиска контактов сотрудников в Active Directory с удобным интерфейсом."/>
                            </BulletDecorator>
                            <BulletDecorator Margin="0,0,0,3">
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="14" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="Реализованы: загрузка данных из AD, текстовый поиск с автодополнением, фильтрация по отделам и должностям."/>
                            </BulletDecorator>
                            <BulletDecorator Margin="0,0,0,3">
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="14" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="Добавлены функции редактирования, экспорта в CSV, печати, автообновления данных и проверки обновлений приложения."/>
                            </BulletDecorator>
                            <BulletDecorator>
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="14" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="Разработан современный дизайн с использованием WPF (Windows Presentation Foundation)."/>
                            </BulletDecorator>
                            <BulletDecorator Margin="0,0,0,3">
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="14" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="v4.0: добавлены тёмная тема, автодополнение поиска, ограничение редактирования по ролям."/>
                            </BulletDecorator>
                            <BulletDecorator>
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="14" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="Разработка ведлась с использованием MiMo Code (Xiaomi) — AI-ассистент модели mimo-auto."/>
                            </BulletDecorator>
                        </StackPanel>

                        <!-- Инструменты и технологии -->
                        <StackPanel Margin="0,0,0,10">
                            <TextBlock Text="🔧  Инструменты и технологии" FontSize="13"
                                       FontWeight="SemiBold" Foreground="#2C3E50" Margin="0,0,0,4"/>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                       Text="PowerShell 5.1+, WPF (XAML), Active Directory (System.DirectoryServices), VS Code, MiMo Code (Xiaomi)"/>
                        </StackPanel>

                        <!-- MiMo Code -->
                        <StackPanel Margin="0,0,0,10">
                            <TextBlock Text="🤖  MiMo Code — AI-ассистент" FontSize="13"
                                       FontWeight="SemiBold" Foreground="#2C3E50" Margin="0,0,0,4"/>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                       Text="MiMo Code — интерактивный CLI-инструмент от Xiaomi для разработки программного обеспечения. Модель mimo-auto используется как основной AI-ассистент на всех этапах lifecycle приложения."/>
                            <BulletDecorator Margin="6,3,0,0">
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="12" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="Анализ архитектуры и кода — рецензирование, выявление проблем безопасности и производительности."/>
                            </BulletDecorator>
                            <BulletDecorator Margin="6,0,0,0">
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="12" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="Разработка функционала — реализация тёмной темы, автодополнения поиска, проверки прав доступа."/>
                            </BulletDecorator>
                            <BulletDecorator Margin="6,0,0,0">
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="12" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="Генерация и редактирование XAML/PowerShell-кода, рефакторинг, тестирование, документирование."/>
                            </BulletDecorator>
                        </StackPanel>

                        <!-- Разработчики и инструменты -->
                        <StackPanel Margin="0,0,0,10">
                            <TextBlock Text="🛠  Разработчики и инструменты" FontSize="13"
                                       FontWeight="SemiBold" Foreground="#2C3E50" Margin="0,0,0,4"/>
                            <BulletDecorator Margin="0,0,0,3">
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="12" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="VS Code — основная среда разработки (редактор кода, отладка, Git-интеграция)."/>
                            </BulletDecorator>
                            <BulletDecorator Margin="0,0,0,3">
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="12" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="Claude Code (Anthropic) — AI-ассистент для генерации и редактирования кода, архитектурных решений, написания тестов и документации."/>
                            </BulletDecorator>
                            <BulletDecorator Margin="0,0,0,3">
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="12" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="DeepSeek — языковая модель, использовавшаяся на начальных этапах разработки для анализа требований и проектирования."/>
                            </BulletDecorator>
                            <BulletDecorator>
                                <BulletDecorator.Bullet>
                                    <TextBlock Text="•" FontSize="12" FontWeight="Bold" Foreground="#5EB048" Margin="0,0,6,0"/>
                                </BulletDecorator.Bullet>
                                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                           Text="MiMo Code (Xiaomi) — AI-ассистент модели mimo-auto, основной инструмент на финальных этапах: рецензирование, доработка, компиляция."/>
                            </BulletDecorator>
                        </StackPanel>

                        <!-- Разделитель -->
                        <Rectangle Height="1" Fill="#ECF0F1" Margin="0,0,0,8"/>

                        <!-- Постановщик задачи -->
                        <StackPanel Margin="0,0,0,4">
                            <TextBlock Text="👤  Постановщик задачи" FontSize="13"
                                       FontWeight="SemiBold" Foreground="#2C3E50" Margin="0,0,0,3"/>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555"
                                       Text="Арбуханов Рустам — формулирование требований, тестирование, организация распространения."/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </DockPanel>
        </Border>
    </Border>
</Window>
'@

# ============================================================
# ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ
# ============================================================
$script:userList = @()
$script:filteredList = @()
$script:isLoading = $false
$script:mainWindow = $null
$script:lastUpdateDate = $null
$script:selectedEmployee = $null
$script:isDomainAdmin = $false
$script:currentUser = ''
$script:refreshTimer = $null
$script:isDarkTheme = $false

# ============================================================
# ПЕРЕКЛЮЧЕНИЕ ТЕМЫ (светлая / тёмная)
# ============================================================
function Set-Theme {
    param([bool]$Dark)
    $script:isDarkTheme = $Dark

    $w = $script:mainWindow
    if ($null -eq $w) { return }

    $palette = if ($Dark) {
        @{ "WindowBgBrush"="#121220"; "CardBgBrush"="#1E1E2E"; "AltRowBgBrush"="#1A1A2A"
           "ColHeaderBgBrush"="#252535"; "PlaceholderBgBrush"="#181828"; "RowBgBrush"="#1E1E2E"
           "InputBgBrush"="#2A2B3E"; "ReadOnlyBgBrush"="#222233"; "HoverBgBrush"="#2D2D42"
           "PressedBgBrush"="#35354A"; "SelectedRowBgBrush"="#2D4A2D"; "CellHoverBgBrush"="#2D4A2D"
           "RowHoverBgBrush"="#252540"; "ChipHoverBgBrush"="#2D4A2D"; "ChipBgBrush"="#2D4A2D"
           "InputBorderBrush"="#3A3B4E"; "DividerBrush"="#333344"; "CardBorderBrush"="#2A2B3E"
           "ReadOnlyBorderBrush"="#333344"; "TextPrimaryBrush"="#D1D5DB"; "TextHeaderBrush"="#E5E7EB"
           "TextSecondaryBrush"="#9CA3AF"; "TextLabelBrush"="#8892A0"; "TextMutedBrush"="#6B7280"
           "TextPlaceholderBrush"="#4B5563"; "TextPlaceholderSubBrush"="#374151"
           "TextAboutBrush"="#9CA3AF"; "TextIconBrush"="#777777" }
    } else {
        @{ "WindowBgBrush"="#F3F6F9"; "CardBgBrush"="#FFFFFF"; "AltRowBgBrush"="#F8FAFC"
           "ColHeaderBgBrush"="#E9ECEF"; "PlaceholderBgBrush"="#FAFBFC"; "RowBgBrush"="#F8FAFC"
           "InputBgBrush"="#FFFFFF"; "ReadOnlyBgBrush"="#F5F5F5"; "HoverBgBrush"="#E8ECF0"
           "PressedBgBrush"="#D5D9DE"; "SelectedRowBgBrush"="#D5EDCC"; "CellHoverBgBrush"="#D5EDCC"
           "RowHoverBgBrush"="#EAF2FA"; "ChipHoverBgBrush"="#D5EDCC"; "ChipBgBrush"="#E8F5D9"
           "InputBorderBrush"="#D0D5DD"; "DividerBrush"="#E0E0E0"; "CardBorderBrush"="#ECF0F1"
           "ReadOnlyBorderBrush"="#E8E8E8"; "TextPrimaryBrush"="#1A1A1A"; "TextHeaderBrush"="#1A1A1A"
           "TextSecondaryBrush"="#666666"; "TextLabelBrush"="#757575"; "TextMutedBrush"="#95A5A6"
           "TextPlaceholderBrush"="#AAAAAA"; "TextPlaceholderSubBrush"="#CCCCCC"
           "TextAboutBrush"="#555555"; "TextIconBrush"="#999999" }
    }

    foreach ($key in $palette.Keys) {
        $oldBrush = $w.TryFindResource($key)
        if ($oldBrush -is [System.Windows.Media.SolidColorBrush]) {
            $newColor = [System.Windows.Media.ColorConverter]::ConvertFromString($palette[$key])
            $newBrush = [System.Windows.Media.SolidColorBrush]::new($newColor)
            $newBrush.Freeze()
            $w.Resources[$key] = $newBrush
        }
    }

    # Обновить кнопку темы
    $themeBtn = $w.FindName("ThemeToggleButton")
    if ($null -ne $themeBtn) {
        $themeBtn.Content = if ($Dark) { "☀️" } else { "🌙" }
        $themeBtn.ToolTip = if ($Dark) { "Светлая тема" } else { "Тёмная тема" }
    }

    # Обновить статус-бар
    $st = $w.FindName("StatusText")
    if ($null -ne $st -and $st.Text -ne '') {
        $themeLabel = if ($Dark) { "Тёмная тема" } else { "Светлая тема" }
        if ($st.Text -notmatch 'Тема:') {
            $st.Text += "  |  Тема: $themeLabel"
        } else {
            $st.Text = $st.Text -replace 'Тема: [^\|]+', "Тема: $themeLabel"
        }
    }
}

# ============================================================
# ПРОВЕРКА ПРАВ ДОМЕННОГО АДМИНИСТРАТОРА
# ============================================================
function Test-DomainAdmin {
    # Способ 1: быстрый — через whoami (не требует AD)
    try {
        $output = whoami /groups 2>$null
        if ($output -match 'Domain Admins' -or $output -match 'Администраторы домена') {
            return $true
        }
    } catch {}

    # Способ 2: через DirectorySearcher (memberOf)
    try {
        $dcServer = "srv-dc-002.e5dag.ru"
        $ldapPath = "LDAP://$dcServer"
        $entry = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
        $searcher.Filter = "(&(sAMAccountName=$([Environment]::UserName)))"
        [void]$searcher.PropertiesToLoad.Add("memberOf")
        $result = $searcher.FindOne()
        if ($null -ne $result) {
            $memberOf = $result.Properties["memberOf"]
            if ($null -ne $memberOf) {
                foreach ($dn in $memberOf) {
                    if ($dn -match 'CN=Domain Admins') { return $true }
                }
            }
        }
        $searcher.Dispose()
        $entry.Dispose()
    } catch {
        Write-Host "[Test-DomainAdmin] AD недоступен" -ForegroundColor Yellow
    }

    return $false
}

# ============================================================
# ЗАГРУЗКА ДАННЫХ ИЗ ACTIVE DIRECTORY
# ============================================================
function Get-ADUsers {
    Write-Progress -Activity "Загрузка AD" -Status "Выполнение запроса..." -PercentComplete 0

    $dcServer = "srv-dc-002.e5dag.ru"
    $entry = $null
    try {
        $entry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$dcServer")
    } catch {
        throw "Не удалось подключиться к $dcServer`: $($_.Exception.Message)"
    }
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
    $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
    [void]$searcher.PropertiesToLoad.AddRange(@("givenName","sn","title","department","telephoneNumber","mail","samAccountName","distinguishedName"))
    $searcher.PageSize = 1000
    $results = $searcher.FindAll()
    $users = New-Object System.Collections.ArrayList
    $total = $results.Count; $i = 0; $excludedCount = 0
    foreach ($r in $results) {
        $i++
        Write-Progress -Activity "Загрузка AD" -Status "Обработка $i из $total" -PercentComplete ([Math]::Round(($i/$total)*100))
        $p = $r.Properties
        $gn = if ($p.givenname) { $p.givenname[0].ToString() } else { '' }
        $sn = if ($p.sn) { $p.sn[0].ToString() } else { '' }
        if ([string]::IsNullOrEmpty($gn) -or [string]::IsNullOrEmpty($sn)) { continue }

        # Исключаем сотрудников ТСД (фамилия или имя начинается на "ТСД")
        if ($sn -like 'ТСД*' -or $gn -like 'ТСД*') { $excludedCount++; continue }

        [void]$users.Add([PSCustomObject]@{
            Surname           = $sn
            GivenName         = $gn
            Title             = if ($p.title) { $p.title[0].ToString() } else { '' }
            Department        = if ($p.department) { $p.department[0].ToString() } else { '' }
            TelephoneNumber   = if ($p.telephonenumber) { $p.telephonenumber[0].ToString() } else { '' }
            Mail              = if ($p.mail) { $p.mail[0].ToString() } else { '' }
            SamAccountName    = if ($p.samaccountname) { $p.samaccountname[0].ToString() } else { '' }
            DistinguishedName = if ($p.distinguishedname) { $p.distinguishedname[0].ToString() } else { '' }
        })
    }
    $results.Dispose(); $searcher.Dispose()
    Write-Progress -Activity "Загрузка AD" -Completed
    return $users.ToArray()
}

# ============================================================
# ПОЛУЧЕНИЕ УНИКАЛЬНЫХ ЗНАЧЕНИЙ ДЛЯ ФИЛЬТРОВ
# ============================================================
function Get-UniqueValues {
    param([string]$Property)
    return $script:userList | Where-Object { $_."$Property" -ne '' -and $_."$Property" -ne '—' } | ForEach-Object { $_."$Property" } | Sort-Object -Unique
}

# ============================================================
# ОБНОВЛЕНИЕ ФИЛЬТРОВ (отдел, должность)
# ============================================================
function Update-FilterLists {
    $df = $script:mainWindow.FindName("DepartmentFilter")
    $tf = $script:mainWindow.FindName("TitleFilter")
    if ($null -eq $df) { return }

    $prevDept = $df.SelectedItem
    $prevTitle = $tf.SelectedItem

    $depts = Get-UniqueValues -Property "Department"
    $titles = Get-UniqueValues -Property "Title"

    $df.Items.Clear(); [void]$df.Items.Add("Все отделы")
    foreach ($d in $depts) { [void]$df.Items.Add($d) }
    if ($null -ne $prevDept -and $df.Items.Contains($prevDept)) { $df.SelectedItem = $prevDept }
    else { $df.SelectedIndex = 0 }

    $tf.Items.Clear(); [void]$tf.Items.Add("Все должности")
    foreach ($t in $titles) { [void]$tf.Items.Add($t) }
    if ($null -ne $prevTitle -and $tf.Items.Contains($prevTitle)) { $tf.SelectedItem = $prevTitle }
    else { $tf.SelectedIndex = 0 }
}

# ============================================================
# ФИЛЬТРАЦИЯ (AND логика: текст + отдел + должность)
# ============================================================
function Update-TableView {
    $sb = $script:mainWindow.FindName("SearchBox")
    $df = $script:mainWindow.FindName("DepartmentFilter")
    $tf = $script:mainWindow.FindName("TitleFilter")
    $dg = $script:mainWindow.FindName("DataGrid")
    $ct = $script:mainWindow.FindName("CountText")
    $tc = $script:mainWindow.FindName("TotalCountText")

    if ($script:userList.Count -eq 0) {
        $dg.ItemsSource = @(); $ct.Text = "Нет данных"
        $script:filteredList = @(); return
    }

    $searchText = $sb.Text.Trim().ToLower()
    $selectedDept = $df.SelectedItem
    $selectedTitle = $tf.SelectedItem

    $result = [System.Collections.ArrayList]::new($script:userList)

    # Фильтр по отделу
    $deptFilter = $null
    if ($null -ne $selectedDept -and $selectedDept -ne 'Все отделы' -and $selectedDept -ne '') {
        $deptFilter = $selectedDept.ToString()
        $result = [System.Collections.ArrayList]@($result | Where-Object { $_.Department -eq $deptFilter })
    }

    # Фильтр по должности
    if ($null -ne $selectedTitle -and $selectedTitle -ne 'Все должности' -and $selectedTitle -ne '') {
        $titleFilter = $selectedTitle.ToString()
        $result = [System.Collections.ArrayList]@($result | Where-Object { $_.Title -eq $titleFilter })
    }

    # Текстовый поиск
    if ($searchText -ne '') {
        $result = [System.Collections.ArrayList]@($result | Where-Object {
            ($_.Surname -like "*$searchText*") -or
            ($_.GivenName -like "*$searchText*") -or
            ($_.Department -like "*$searchText*") -or
            ($_.Title -like "*$searchText*") -or
            ($_.TelephoneNumber -like "*$searchText*") -or
            ($_.Mail -like "*$searchText*")
        })
    }

    $script:filteredList = $result
    $dg.ItemsSource = $result
    $ct.Text = "Найдено: $($result.Count)  |  Всего: $($script:userList.Count)"
    $tc.Text = "Записей: $($script:userList.Count)"
}

# ============================================================
# ОБНОВЛЕНИЕ ПАНЕЛИ БЫСТРОГО ПРОСМОТРА
# ============================================================
function Update-PreviewPanel {
    param($Employee)

    $placeholder = $script:mainWindow.FindName("PreviewPlaceholder")
    $avatarBorder = $script:mainWindow.FindName("AvatarBorder")
    $avatarText   = $script:mainWindow.FindName("AvatarText")
    $previewName  = $script:mainWindow.FindName("PreviewName")
    $previewTitleRow = $script:mainWindow.FindName("PreviewTitleRow")
    $previewTitle  = $script:mainWindow.FindName("PreviewTitle")
    $previewDeptRow = $script:mainWindow.FindName("PreviewDeptRow")
    $previewDept   = $script:mainWindow.FindName("PreviewDept")
    $previewDivider = $script:mainWindow.FindName("PreviewDivider")
    $previewPhoneRow = $script:mainWindow.FindName("PreviewPhoneRow")
    $previewPhone  = $script:mainWindow.FindName("PreviewPhone")
    $previewEmailRow = $script:mainWindow.FindName("PreviewEmailRow")
    $previewEmail  = $script:mainWindow.FindName("PreviewEmail")
    $previewEdit   = $script:mainWindow.FindName("PreviewEditButton")

    if ($null -eq $Employee) {
        $placeholder.Visibility = "Visible"
        $avatarBorder.Visibility = "Collapsed"
        $previewName.Visibility = "Collapsed"
        $previewTitleRow.Visibility = "Collapsed"
        $previewDeptRow.Visibility = "Collapsed"
        $previewDivider.Visibility = "Collapsed"
        $previewPhoneRow.Visibility = "Collapsed"
        $previewEmailRow.Visibility = "Collapsed"
        $previewEdit.Visibility = "Collapsed"
        return
    }

    $script:selectedEmployee = $Employee
    $placeholder.Visibility = "Collapsed"

    # Формируем инициалы
    $initials = "$($Employee.Surname[0])$($Employee.GivenName[0])"
    $fullName = "$($Employee.Surname) $($Employee.GivenName)".Trim()

    $avatarText.Text = $initials
    $avatarBorder.Visibility = "Visible"

    $previewName.Text = $fullName
    $previewName.Visibility = "Visible"

    if ($Employee.Title -ne '' -and $Employee.Title -ne '—') {
        $previewTitle.Text = $Employee.Title
        $previewTitleRow.Visibility = "Visible"
    } else { $previewTitleRow.Visibility = "Collapsed" }

    if ($Employee.Department -ne '' -and $Employee.Department -ne '—') {
        $previewDept.Text = $Employee.Department
        $previewDeptRow.Visibility = "Visible"
    } else { $previewDeptRow.Visibility = "Collapsed" }

    if ($Employee.TelephoneNumber -ne '' -and $Employee.TelephoneNumber -ne '—') {
        $previewPhone.Text = $Employee.TelephoneNumber
        $previewPhoneRow.Visibility = "Visible"
    } else { $previewPhoneRow.Visibility = "Collapsed" }

    if ($Employee.Mail -ne '' -and $Employee.Mail -ne '—') {
        $previewEmail.Text = $Employee.Mail
        $previewEmailRow.Visibility = "Visible"
    } else { $previewEmailRow.Visibility = "Collapsed" }

    $previewDivider.Visibility = "Visible"
    $previewEdit.Visibility = "Visible"
}

# ============================================================
# УТИЛИТА: Проверка пароля доменного администратора через LDAP
# ============================================================
function Test-AdminPassword {
    param([string]$Login, [string]$Password)
    if ([string]::IsNullOrEmpty($Login) -or [string]::IsNullOrEmpty($Password)) { return $false }
    try {
        $dcServer = 'srv-dc-002.e5dag.ru'
        $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext(
            [System.DirectoryServices.AccountManagement.ContextType]::Domain, $dcServer
        )
        $valid = $pc.ValidateCredentials($Login, $Password)
        $pc.Dispose()
        if (-not $valid) { return $false }

        # Проверка группы: whoami (локально)
        $whoamiOutput = whoami /groups 2>$null
        if ($whoamiOutput -match 'Domain Admins' -or $whoamiOutput -match 'Администраторы домена') {
            return $true
        }

        # Fallback: memberOf через DirectorySearcher
        $ldapPath = "LDAP://$dcServer"
        $entry = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
        $searcher.Filter = "(&(sAMAccountName=$Login))"
        [void]$searcher.PropertiesToLoad.Add("memberOf")
        $result = $searcher.FindOne()
        if ($null -ne $result) {
            $memberOf = $result.Properties["memberOf"]
            if ($null -ne $memberOf) {
                foreach ($dn in $memberOf) {
                    if ($dn -match 'CN=Domain Admins') { return $true }
                }
            }
        }
        $searcher.Dispose()
        $entry.Dispose()
        return $false
    } catch {
        return $false
    }
}

# ============================================================
# ДИАЛОГ ВХОДА (обязательная аутентификация при запуске)
# Returns: $true если вход успешен, $false если отмена/ошибка
# ============================================================
function Show-LoginDialog {
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xamlLoginDialog)
        $lw = [System.Windows.Markup.XamlReader]::Load($reader)
        $reader.Dispose()

        $loginBox    = $lw.FindName('LoginBox')
        $passwordBox = $lw.FindName('LoginPasswordBox')
        $errorText   = $lw.FindName('LoginErrorText')
        $statusText  = $lw.FindName('LoginStatusText')
        $okBtn       = $lw.FindName('LoginOkButton')
        $cancelBtn   = $lw.FindName('LoginCancelButton')

        $loginBox.Text = [Environment]::UserName
        $statusText.Text = "Введите логин и пароль доменного администратора."

        $script:loginResult = $false

        $okBtn.Add_Click({
            $login = $loginBox.Text.Trim()
            $pass  = $passwordBox.Password

            if ([string]::IsNullOrEmpty($login) -or [string]::IsNullOrEmpty($pass)) {
                $errorText.Text = "Заполните логин и пароль."
                $errorText.Visibility = "Visible"
                return
            }

            $okBtn.IsEnabled = $false
            $okBtn.Content = "⏳ Проверка..."
            $errorText.Visibility = "Collapsed"
            $statusText.Text = "Проверка учётных данных..."

            try {
                $pc = $null
                $dcServer = 'srv-dc-002.e5dag.ru'

                $statusText.Text = "Подключение к контроллеру домена..."
                try {
                    $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext(
                        [System.DirectoryServices.AccountManagement.ContextType]::Domain, $dcServer
                    )
                } catch {
                    $errMsg = $_.Exception.Message
                    $errorText.Text = "Не удалось подключиться к $dcServer`: $errMsg"
                    $errorText.Visibility = "Visible"
                    $okBtn.IsEnabled = $true
                    $okBtn.Content = "✔  Войти"
                    $statusText.Text = ""
                    return
                }

                # Шаг 2: Валидация учётных данных
                $statusText.Text = "Проверка учётных данных..."
                $valid = $pc.ValidateCredentials($login, $pass)

                if (-not $valid) {
                    $errorText.Text = "Неверный логин или пароль."
                    $errorText.Visibility = "Visible"
                    $okBtn.IsEnabled = $true
                    $okBtn.Content = "✔  Войти"
                    $statusText.Text = ""
                    $passwordBox.Password = ""
                    $passwordBox.Focus() | Out-Null
                    $pc.Dispose()
                    return
                }

                # Шаг 3: Проверка группы Domain Admins (как в Export-Phonebook.ps1)
                $statusText.Text = "Проверка прав доступа..."
                $pc.Dispose()

                # Способ 1: whoami /groups (локально, без LDAP)
                $whoamiOutput = whoami /groups 2>$null
                $isAdmin = $false
                if ($whoamiOutput -match 'Domain Admins' -or $whoamiOutput -match 'Администраторы домена') {
                    $isAdmin = $true
                }

                # Способ 2: Если whoami не сработал — через DirectorySearcher (memberOf)
                if (-not $isAdmin) {
                    try {
                        $ldapPath = "LDAP://$dcServer"
                        $entry = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
                        $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
                        $searcher.Filter = "(&(sAMAccountName=$login))"
                        [void]$searcher.PropertiesToLoad.Add("memberOf")
                        $result = $searcher.FindOne()
                        if ($null -ne $result) {
                            $memberOf = $result.Properties["memberOf"]
                            if ($null -ne $memberOf) {
                                foreach ($dn in $memberOf) {
                                    if ($dn -match 'CN=Domain Admins') { $isAdmin = $true; break }
                                }
                            }
                        }
                        $searcher.Dispose()
                        $entry.Dispose()
                    } catch {}
                }

                if (-not $isAdmin) {
                    $errorText.Text = "Доступ запрещён. Требуется членство в группе «Domain Admins»."
                    $errorText.Visibility = "Visible"
                    $okBtn.IsEnabled = $true
                    $okBtn.Content = "✔  Войти"
                    $statusText.Text = ""
                    $passwordBox.Password = ""
                    $passwordBox.Focus() | Out-Null
                    return
                }

                # Успех
                $script:isDomainAdmin = $true
                $script:currentUser = $login
                $script:loginResult = $true
                $lw.Close()

            } catch {
                $errMsg = $_.Exception.Message
                $errorText.Text = "Ошибка: $errMsg"
                $errorText.Visibility = "Visible"
                $okBtn.IsEnabled = $true
                $okBtn.Content = "✔  Войти"
                $statusText.Text = ""
            }
        })

        $cancelBtn.Add_Click({ $lw.Close() })

        $passwordBox.Add_KeyDown({
            if ($_.Key -eq 'Enter') {
                $okBtn.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            }
        })

        [void]$lw.ShowDialog()
        return $script:loginResult

    } catch {
        Write-Host "[Show-LoginDialog] Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# ============================================================
# ДИАЛОГ ПАРОЛЯ АДМИНИСТРАТОРА
# Returns: $true если пароль верный, $false если отмена
# ============================================================
function Show-AdminPasswordDialog {
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xamlPasswordDialog)
        $pw = [System.Windows.Markup.XamlReader]::Load($reader)
        $reader.Dispose()

        $loginBox    = $pw.FindName('AdminLoginBox')
        $passwordBox = $pw.FindName('AdminPasswordBox')
        $errorText   = $pw.FindName('AuthErrorText')
        $statusText  = $pw.FindName('AuthStatusText')
        $okBtn       = $pw.FindName('AuthOkButton')
        $cancelBtn   = $pw.FindName('AuthCancelButton')

        # Pre-fill login with current user
        $loginBox.Text = [Environment]::UserName
        $statusText.Text = "Введите пароль доменного администратора для разблокировки редактирования."

        $script:adminDialogResult = $false

        $okBtn.Add_Click({
            $login = $loginBox.Text.Trim()
            $pass  = $passwordBox.Password

            if ([string]::IsNullOrEmpty($pass)) {
                $errorText.Text = "Введите пароль."
                $errorText.Visibility = "Visible"
                return
            }

            $okBtn.IsEnabled = $false
            $okBtn.Content = "⏳ Проверка..."
            $errorText.Visibility = "Collapsed"
            $statusText.Text = "Проверка учётных данных..."

            if (Test-AdminPassword -Login $login -Password $pass) {
                $script:isEditUnlocked = $true
                $script:adminDialogResult = $true
                $pw.Close()
            } else {
                $errorText.Text = "Неверный логин или пароль, либо вы не являетесь доменным администратором."
                $errorText.Visibility = "Visible"
                $okBtn.IsEnabled = $true
                $okBtn.Content = "✔  Войти"
                $statusText.Text = ""
                $passwordBox.Password = ""
                $passwordBox.Focus() | Out-Null
            }
        })

        $cancelBtn.Add_Click({ $pw.Close() })

        # Enter key triggers OK
        $passwordBox.Add_KeyDown({
            if ($_.Key -eq 'Enter') { $okBtn.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }
        })

        $pw.Owner = $script:mainWindow
        [void]$pw.ShowDialog()

        return $script:adminDialogResult
    } catch {
        Write-Host "[Show-AdminPasswordDialog] Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# ============================================================
# УТИЛИТА: Запись атрибута в AD (с очисткой при пустом значении)
# ============================================================
function Set-ADAttribute {
    param(
        [Parameter(Mandatory)] $Entry,
        [Parameter(Mandatory)][string] $Attribute,
        [string] $Value
    )
    if (-not [string]::IsNullOrEmpty($Value)) {
        $Entry.Properties[$Attribute].Value = $Value
    } else {
        if ($Entry.Properties[$Attribute].Count -gt 0) {
            $Entry.Properties[$Attribute].Clear()
        }
    }
}

# ============================================================
# УТИЛИТА: Санитизация строки (удаление LDAP-опасных символов)
# ============================================================
function Sanitize-ADString {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    # Удаляем управляющие символы + LDAP-опасные: , = + < > # \ "
    return ($Value -replace '[\x00-\x1f]', '' -replace '[,=+<>#\\"]', '')
}

# ============================================================
# УТИЛИТА: Безопасное сообщение об ошибке AD
# ============================================================
function Show-ADErrorMessage {
    param([string]$Context, [string]$ExceptionMessage)
    [System.Windows.MessageBox]::Show(
        "Ошибка [$Context]: $ExceptionMessage",
        "Ошибка подключения",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
}

# ============================================================
# КАРТОЧКА СОТРУДНИКА (модальное окно) — ИСПРАВЛЕННАЯ
# ============================================================
function Show-EmployeeCard {
    param($Employee, [bool]$EditMode = $false)

    # Проверка: сотрудник должен быть выбран
    if ($null -eq $Employee) {
        Write-Host "[Show-EmployeeCard] ОШИБКА: Employee = null" -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Не выбран сотрудник.", "Ошибка",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # Проверка прав при попытке редактирования
    if ($EditMode -and -not $script:isDomainAdmin) {
        [System.Windows.MessageBox]::Show(
            "Только доменные администраторы могут редактировать данные.",
            "Доступ запрещён",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }

    try {
        # Загружаем XAML
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xamlCardWindow)
        $cw = [System.Windows.Markup.XamlReader]::Load($reader)
        $reader.Dispose()

        # ----------------------------------------------------------
        # Поиск всех элементов управления по имени
        # ----------------------------------------------------------
        $ct   = $cw.FindName("CardTitle")        # Заголовок
        $gn   = $cw.FindName("GivenNameBox")     # Имя
        $sn   = $cw.FindName("SurnameBox")       # Фамилия
        $ti   = $cw.FindName("TitleBox")         # Должность
        $dp   = $cw.FindName("DepartmentBox")    # Отдел
        $ph   = $cw.FindName("TelephoneBox")     # Телефон
        $em   = $cw.FindName("MailBox")          # Email
        $lg   = $cw.FindName("LoginBox")         # Логин
        $btnCopy = $cw.FindName("CopyPhoneButton")  # Кнопка "Копировать телефон"
        $btnSave = $cw.FindName("SaveButton")        # Кнопка "Сохранить"
        $btnClose = $cw.FindName("CloseButton")      # Кнопка "Закрыть"

        # ----------------------------------------------------------
        # Проверка: все ли элементы найдены
        # ----------------------------------------------------------
        $allFound = $true
        @{ "CardTitle"=$ct; "GivenNameBox"=$gn; "SurnameBox"=$sn; "TitleBox"=$ti;
           "DepartmentBox"=$dp; "TelephoneBox"=$ph; "MailBox"=$em; "LoginBox"=$lg;
           "CopyPhoneButton"=$btnCopy; "SaveButton"=$btnSave; "CloseButton"=$btnClose
        }.GetEnumerator() | ForEach-Object {
            if ($null -eq $_.Value) {
                Write-Host "[Show-EmployeeCard] ПРЕДУПРЕЖДЕНИЕ: не найден элемент '$($_.Key)'" -ForegroundColor Yellow
                $allFound = $false
            }
        }

        # ----------------------------------------------------------
        # Заполняем поля данными сотрудника
        # ----------------------------------------------------------
        $gn.Text  = $Employee.GivenName
        $sn.Text  = $Employee.Surname
        # Замена "—" на пустые строки для редактирования
        $ti.Text  = if ($Employee.Title -eq '—') { '' } else { $Employee.Title }
        $dp.Text  = if ($Employee.Department -eq '—') { '' } else { $Employee.Department }
        $ph.Text  = if ($Employee.TelephoneNumber -eq '—') { '' } else { $Employee.TelephoneNumber }
        $em.Text  = if ($Employee.Mail -eq '—') { '' } else { $Employee.Mail }
        $lg.Text  = $Employee.SamAccountName
        $display  = "$($Employee.Surname) $($Employee.GivenName)".Trim()

        # ----------------------------------------------------------
        # Настройка режима: редактирование или просмотр
        # ----------------------------------------------------------
        if ($EditMode) {
            # === РЕЖИМ РЕДАКТИРОВАНИЯ ===
            $cw.Title   = "Редактирование: $display"
            $ct.Text    = "✏️  Редактирование: $display"

            # Показываем кнопку "Сохранить"
            if ($null -ne $btnSave) { $btnSave.Visibility = "Visible" }

            # Доступны для редактирования только: Должность, Отдел, Телефон, Почта
            # Имя, Фамилия, Логин — только чтение (ReadOnlyBox в XAML)

            # Обработчик кнопки "Сохранить"
            if ($null -ne $btnSave) {
                $btnSave.Add_Click({
                    $btnSave.IsEnabled = $false
                    $btnSave.Content   = "⏳ Сохранение..."

                    # Собираем данные из доступных для редактирования полей
                    $newTI  = $ti.Text.Trim()
                    $newDP  = $dp.Text.Trim()
                    $newPH  = $ph.Text.Trim()
                    $newEM  = $em.Text.Trim()

                    # === ВАЛИДАЦИЯ ВХОДНЫХ ДАННЫХ ===
                    # Ограничение длины
                    if ($newTI.Length -gt 64) {
                        [System.Windows.MessageBox]::Show("Должность: максимум 64 символа.", "Ошибка",
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        $btnSave.IsEnabled = $true; $btnSave.Content = "💾  Сохранить"; return
                    }
                    if ($newDP.Length -gt 64) {
                        [System.Windows.MessageBox]::Show("Отдел: максимум 64 символа.", "Ошибка",
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        $btnSave.IsEnabled = $true; $btnSave.Content = "💾  Сохранить"; return
                    }
                    if ($newPH.Length -gt 20) {
                        [System.Windows.MessageBox]::Show("Телефон: максимум 20 символов.", "Ошибка",
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        $btnSave.IsEnabled = $true; $btnSave.Content = "💾  Сохранить"; return
                    }
                    if ($newEM.Length -gt 256) {
                        [System.Windows.MessageBox]::Show("Email: максимум 256 символов.", "Ошибка",
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        $btnSave.IsEnabled = $true; $btnSave.Content = "💾  Сохранить"; return
                    }
                    # Валидация email формата (если не пустой)
                    if (-not [string]::IsNullOrEmpty($newEM) -and $newEM -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                        [System.Windows.MessageBox]::Show("Некорректный формат email.", "Ошибка",
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        $btnSave.IsEnabled = $true; $btnSave.Content = "💾  Сохранить"; return
                    }
                    # Фильтрация LDAP-опасных символов
                    $sanitizedTI = $newTI -replace '[\x00-\x1f]', ''
                    $sanitizedDP = $newDP -replace '[\x00-\x1f]', ''
                    $sanitizedPH = $newPH -replace '[\x00-\x1f]', ''
                    $sanitizedEM = $newEM -replace '[\x00-\x1f]', ''

                    # Сохраняем изменения в Active Directory
                    $entry = $null
                    try {
                        $entry = New-Object System.DirectoryServices.DirectoryEntry(
                            "LDAP://$($Employee.DistinguishedName)"
                        )
                        $null = $entry.NativeObject

                        Set-ADAttribute -Entry $entry -Attribute "title"            -Value $sanitizedTI
                        Set-ADAttribute -Entry $entry -Attribute "department"       -Value $sanitizedDP
                        Set-ADAttribute -Entry $entry -Attribute "telephoneNumber"  -Value $sanitizedPH
                        Set-ADAttribute -Entry $entry -Attribute "mail"             -Value $sanitizedEM

                        $entry.CommitChanges()

                        $Employee.Title            = $sanitizedTI
                        $Employee.Department       = $sanitizedDP
                        $Employee.TelephoneNumber  = $sanitizedPH
                        $Employee.Mail             = $sanitizedEM

                        Update-FilterLists
                        Update-TableView
                        Update-PreviewPanel -Employee $Employee

                        [System.Windows.MessageBox]::Show(
                            "Данные сотрудника сохранены в Active Directory.",
                            "Готово",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Information
                        )
                        $cw.Close()
                    } catch {
                        # Санитизация: не показываем внутренние детали AD
                        Write-Host "[Edit] Ошибка: $($_.Exception.Message)" -ForegroundColor Red
                        [System.Windows.MessageBox]::Show(
                            "Не удалось сохранить данные. Обратитесь к администратору.",
                            "Ошибка",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Error
                        )
                        $btnSave.IsEnabled = $true
                        $btnSave.Content   = "💾  Сохранить"
                    } finally {
                        if ($null -ne $entry) { $entry.Dispose() }
                    }
                })
            }
        } else {
            # === РЕЖИМ ПРОСМОТРА (только чтение) ===
            $cw.Title   = "Просмотр: $display"
            $ct.Text    = "👤  $display"

            # Кнопка "Сохранить" остаётся скрытой (в XAML Visibility="Collapsed")
            # Блокируем поля, которые могут быть редактируемыми (Должность, Отдел, Телефон, Почта)
            # Имя, Фамилия, Логин — уже ReadOnlyBox в XAML
            $readOnlyStyle = $cw.TryFindResource("ReadOnlyBox")
            if ($null -ne $readOnlyStyle) {
                $ti.Style = $readOnlyStyle
                $dp.Style = $readOnlyStyle
                $ph.Style = $readOnlyStyle
                $em.Style = $readOnlyStyle
            }
        }

        # ----------------------------------------------------------
        # ОБЩИЕ ОБРАБОТЧИКИ (для обоих режимов)
        # ----------------------------------------------------------

        # Кнопка "Копировать телефон"
        if ($null -ne $btnCopy) {
            $btnCopy.Add_Click({
                $phoneBox = $cw.FindName("TelephoneBox")
                $phoneText = if ($null -ne $phoneBox) { $phoneBox.Text.Trim() } else { '' }
                if ($phoneText -ne '') {
                    [System.Windows.Clipboard]::SetText($phoneText)
                    $originalContent = $btnCopy.Content
                    $btnCopy.Content   = "✅ Скопировано!"
                    $btnCopy.IsEnabled = $false
                    $timer = New-Object System.Windows.Threading.DispatcherTimer
                    $timer.Interval = [TimeSpan]::FromMilliseconds(1500)
                    $timer.Add_Tick({
                        $timer.Stop()
                        $btnCopy.Content   = $originalContent
                        $btnCopy.IsEnabled = $true
                    })
                    $timer.Start()
                } else {
                    [System.Windows.MessageBox]::Show(
                        "Телефон не указан.",
                        "Информация",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )
                }
            })
        }

        # Кнопка "Закрыть"
        if ($null -ne $btnClose) {
            $btnClose.Add_Click({ $cw.Close() })
        }

        # ----------------------------------------------------------
        # Открываем окно (модально, поверх главного)
        # ----------------------------------------------------------
        $cw.Owner = $script:mainWindow
        [void]$cw.ShowDialog()

    } catch {
        Show-ADErrorMessage -Context "EmployeeCard" -ExceptionMessage $_.Exception.Message
    }
}

# ============================================================
# СОЗДАНИЕ НОВОГО СОТРУДНИКА
# ============================================================
function Show-CreateContactCard {
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xamlCardWindow)
        $cw = [System.Windows.Markup.XamlReader]::Load($reader)
        $reader.Dispose()

        $ct   = $cw.FindName("CardTitle")
        $gn   = $cw.FindName("GivenNameBox")
        $sn   = $cw.FindName("SurnameBox")
        $ti   = $cw.FindName("TitleBox")
        $dp   = $cw.FindName("DepartmentBox")
        $ph   = $cw.FindName("TelephoneBox")
        $em   = $cw.FindName("MailBox")
        $lg   = $cw.FindName("LoginBox")
        $btnCopy = $cw.FindName("CopyPhoneButton")
        $btnSave = $cw.FindName("SaveButton")
        $btnClose = $cw.FindName("CloseButton")

        $cw.Title = "Новый сотрудник"
        $ct.Text  = "＋ Новый сотрудник"

        # Все поля редактируемые (включая Имя и Фамилию)
        $editStyle = $cw.TryFindResource("FieldBox")
        if ($null -ne $editStyle) {
            $gn.Style = $editStyle
            $sn.Style = $editStyle
        }

        # Скрыть кнопку копирования телефона (некопировать)
        if ($null -ne $btnCopy) { $btnCopy.Visibility = "Collapsed" }

        # Показать кнопку "Сохранить"
        if ($null -ne $btnSave) { $btnSave.Visibility = "Visible" }

        # Обработчик "Сохранить"
        if ($null -ne $btnSave) {
            $btnSave.Add_Click({
                $newGN = $gn.Text.Trim()
                $newSN = $sn.Text.Trim()
                $newTI = $ti.Text.Trim()
                $newDP = $dp.Text.Trim()
                $newPH = $ph.Text.Trim()
                $newEM = $em.Text.Trim()

                # === ВАЛИДАЦИЯ ===
                if ([string]::IsNullOrEmpty($newGN) -or [string]::IsNullOrEmpty($newSN)) {
                    [System.Windows.MessageBox]::Show("Имя и фамилия обязательны.","Ошибка",
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                # Допустимые символы: буквы (кириллица/латиница), пробел, дефис, апостроф
                if ($newGN -notmatch "^[а-яА-ЯёЁa-zA-Z\s\-']+$" -or $newSN -notmatch "^[а-яА-ЯёЁa-zA-Z\s\-']+$") {
                    [System.Windows.MessageBox]::Show("Имя и фамилия: допустимы только буквы, пробел, дефис.","Ошибка",
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                # Валидация email (если указан)
                if (-not [string]::IsNullOrEmpty($newEM) -and $newEM -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                    [System.Windows.MessageBox]::Show("Некорректный формат email.","Ошибка",
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                # Ограничение длины
                if ($newTI.Length -gt 64 -or $newDP.Length -gt 64 -or $newPH.Length -gt 20 -or $newEM.Length -gt 256) {
                    [System.Windows.MessageBox]::Show("Превышена допустимая длина поля.","Ошибка",
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }

                # Санитизация всех строк
                $safeGN = Sanitize-ADString $newGN
                $safeSN = Sanitize-ADString $newSN
                $safeTI = $newTI -replace '[\x00-\x1f]', ''
                $safeDP = $newDP -replace '[\x00-\x1f]', ''
                $safePH = $newPH -replace '[\x00-\x1f]', ''
                $safeEM = $newEM -replace '[\x00-\x1f]', ''

                # Генерируем sAMAccountName (только ASCII, макс. 20 символов)
                $samBase = "$($safeGN[0])$($safeSN)".ToLower() -replace '[^a-z0-9]',''
                if ([string]::IsNullOrEmpty($samBase)) {
                    [System.Windows.MessageBox]::Show("Не удалось сгенерировать логин. Проверьте имя и фамилию.","Ошибка",
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                if ($samBase.Length -gt 20) { $samBase = $samBase.Substring(0, 20) }

                # Формируем CN безопасно (экранируем спецсимволы LDAP)
                $safeCN = "$safeSN $safeGN" -replace '[,=+<>#\\"]', ''

                # Генерируем временный пароль (случайный, 16 символов)
                $symbols = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%&*'
                $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
                $bytes = New-Object byte[] 16
                $rng.GetBytes($bytes)
                $tempPass = -join ($bytes | ForEach-Object { $symbols[$_ % $symbols.Length] })
                $securePass = ConvertTo-SecureString $tempPass -AsPlainText -Force

                $entry = $null
                $newUser = $null
                try {
                    $entry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://DC=e5dag,DC=ru")
                    $newUser = $entry.Children.Add("CN=$safeCN", "user")
                    $newUser.Properties["sAMAccountName"].Value = $samBase
                    $newUser.Properties["givenName"].Value = $safeGN
                    $newUser.Properties["sn"].Value = $safeSN
                    if (-not [string]::IsNullOrEmpty($safeTI)) { $newUser.Properties["title"].Value = $safeTI }
                    if (-not [string]::IsNullOrEmpty($safeDP)) { $newUser.Properties["department"].Value = $safeDP }
                    if (-not [string]::IsNullOrEmpty($safePH)) { $newUser.Properties["telephoneNumber"].Value = $safePH }
                    if (-not [string]::IsNullOrEmpty($safeEM)) { $newUser.Properties["mail"].Value = $safeEM }

                    # Устанавливаем пароль и активируем учётную запись
                    $newUser.Invoke("SetPassword", $tempPass)
                    $newUser.Properties["userAccountControl"].Value = 0x0200  # Normal account
                    $newUser.CommitChanges()

                    # Временный пароль не показываем — он хранится в AD (сброс при первом входе)
                    [System.Windows.MessageBox]::Show(
                        "Сотрудник '$safeSN $safeGN' создан в Active Directory.`nЛогин: $samBase`nПароль: установлен (сброс при первом входе)",
                        "Готово", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

                    Update-FilterLists
                    Update-TableView
                    $cw.Close()
                } catch {
                    Show-ADErrorMessage -Context "CreateUser" -ExceptionMessage $_.Exception.Message
                } finally {
                    if ($null -ne $newUser) { $newUser.Dispose() }
                    if ($null -ne $entry) { $entry.Dispose() }
                }
            })
        }

        # Кнопка "Закрыть"
        if ($null -ne $btnClose) {
            $btnClose.Add_Click({ $cw.Close() })
        }

        $cw.Owner = $script:mainWindow
        [void]$cw.ShowDialog()
    } catch {
        Show-ADErrorMessage -Context "CreateContactCard" -ExceptionMessage $_.Exception.Message
    }
}

# ============================================================
# ОКНО «О ПРОГРАММЕ»
# ============================================================
function Show-AboutWindow {
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xamlAboutWindow)
        $aw = [System.Windows.Markup.XamlReader]::Load($reader)
        $reader.Dispose()

        # Устанавливаем версию (из .exe или по умолчанию)
        $versionText = $aw.FindName("VersionText")
        if ($null -ne $versionText) {
            $versionText.Text = $script:appVersion
        }

        # Обработчик кнопки "Закрыть"
        $closeBtn = $aw.FindName("AboutCloseButton")
        if ($null -ne $closeBtn) {
            $closeBtn.Add_Click({ $aw.Close() })
        }

        $aw.Owner = $script:mainWindow
        [void]$aw.ShowDialog()
    } catch {
        Show-ADErrorMessage -Context "AboutWindow" -ExceptionMessage $_.Exception.Message
    }
}

# ============================================================
# ЭКСПОРТ В CSV
# ============================================================
function Export-ToCsv {
    if ($script:filteredList.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Нет данных для экспорта.","Нет данных",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Warning)
        return
    }
    $sd = New-Object Microsoft.Win32.SaveFileDialog
    $sd.Title = "Экспорт телефонного справочника"
    $sd.Filter = "CSV (UTF-8) (*.csv)|*.csv|Все файлы (*.*)|*.*"
    $sd.DefaultExt = ".csv"
    $sd.FileName = "Phonebook_$(Get-Date -Format 'yyyy-MM-dd').csv"
    $sd.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if ($sd.ShowDialog() -eq $true) {
        $data = foreach ($e in $script:filteredList) {
            # Замена "—" на пустые строки для экспорта
            $dept = if ($e.Department -eq '—') { '' } else { $e.Department }
            $title = if ($e.Title -eq '—') { '' } else { $e.Title }
            $phone = if ($e.TelephoneNumber -eq '—') { '' } else { $e.TelephoneNumber }
            $mail = if ($e.Mail -eq '—') { '' } else { $e.Mail }
            [PSCustomObject]@{
                "Фамилия"=$e.Surname; "Имя"=$e.GivenName; "Отдел"=$dept
                "Должность"=$title; "Телефон"=$phone; "Email"=$mail; "Логин"=$e.SamAccountName
            }
        }
        $data | Export-Csv -Path $sd.FileName -Delimiter ';' -Encoding UTF8 -NoTypeInformation
        [System.Windows.MessageBox]::Show("Экспортировано: $($data.Count) строк.","Готово",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information)
    }
}

# ============================================================
# ПЕЧАТЬ
# ============================================================
function Invoke-Print {
    if ($script:userList.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Нет данных для печати.","Нет данных",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Warning); return
    }
    $pd = New-Object System.Windows.Controls.PrintDialog
    if ($pd.ShowDialog() -eq $true) {
        $printDate = Get-Date -Format "dd.MM.yyyy HH:mm"
        $printDg = $script:mainWindow.FindName("DataGrid")
        $printDg.PrintVisual($printDg, "Телефонный справочник | $printDate")
    }
}

# ============================================================
# ЗАГРУЗКА ДАННЫХ
# ============================================================
function Invoke-Refresh {
    if ($script:isLoading) { return }
    $script:isLoading = $true

    $sb=$script:mainWindow.FindName("SearchBox"); $df=$script:mainWindow.FindName("DepartmentFilter")
    $tf=$script:mainWindow.FindName("TitleFilter"); $dg=$script:mainWindow.FindName("DataGrid")
    $eb=$script:mainWindow.FindName("ExportButton")
    $st=$script:mainWindow.FindName("StatusText")
    $ct=$script:mainWindow.FindName("CountText"); $si=$script:mainWindow.FindName("StatusIndicator")
    $lu=$script:mainWindow.FindName("LastUpdateText")
    $cu=$script:mainWindow.FindName("CurrentUserText")

    # Отображаем текущего пользователя (после аутентификации)
    if ($null -ne $cu) {
        $domain = [Environment]::UserDomainName
        $cu.Text = "👤 $domain\$($script:currentUser)"
    }

    $sb.IsEnabled=$false; $df.IsEnabled=$false; $tf.IsEnabled=$false
    $eb.IsEnabled=$false
    $dg.ItemsSource=@()
    $si.Fill="#BDC3C7"; $st.Text="Загрузка данных..."; $ct.Text="Загрузка..."
    Update-PreviewPanel -Employee $null
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[System.Windows.Threading.DispatcherPriority]::Render)

    try {
        $script:userList = @(Get-ADUsers)

        # Замена пустых значений на "—" для отображения в таблице
        foreach ($u in $script:userList) {
            if ([string]::IsNullOrEmpty($u.Title))           { $u.Title = "—" }
            if ([string]::IsNullOrEmpty($u.Department))      { $u.Department = "—" }
            if ([string]::IsNullOrEmpty($u.TelephoneNumber)) { $u.TelephoneNumber = "—" }
            if ([string]::IsNullOrEmpty($u.Mail))            { $u.Mail = "—" }
        }

        Update-FilterLists
        $sb.Text = ''
        if ($df.Items.Count -gt 0) { $df.SelectedIndex = 0 }
        if ($tf.Items.Count -gt 0) { $tf.SelectedIndex = 0 }
        Update-TableView
        $si.Fill = "#27AE60"
        $st.Text = "Подключено к AD"
        $script:lastUpdateDate = Get-Date
        $lu.Text = "Данные от $($script:lastUpdateDate.ToString('dd.MM.yyyy HH:mm'))"
        $st.Text = "Подключено к AD  |  $($script:userList.Count) сотрудников  |  Пользователь: $($script:currentUser)"
    } catch {
        $si.Fill = "#C0392B"
        $st.Text = "Ошибка: $($_.Exception.Message)"
        Show-ADErrorMessage -Context "Refresh" -ExceptionMessage $_.Exception.Message
    }

    $sb.IsEnabled=$true; $df.IsEnabled=$true; $tf.IsEnabled=$true
    $eb.IsEnabled=$true
    $script:isLoading = $false
}

# ============================================================
# ЗАПУСК ПРИЛОЖЕНИЯ
# ============================================================
try {
    $xr = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xamlMainWindow)
    $script:mainWindow = [System.Windows.Markup.XamlReader]::Load($xr); $xr.Dispose()

    $sb=$script:mainWindow.FindName("SearchBox"); $cb=$script:mainWindow.FindName("ClearButton")
    $sp=$script:mainWindow.FindName("SearchPlaceholder")
    $df=$script:mainWindow.FindName("DepartmentFilter"); $tf=$script:mainWindow.FindName("TitleFilter")
    $dg=$script:mainWindow.FindName("DataGrid"); $eb=$script:mainWindow.FindName("ExportButton")
    $crb=$script:mainWindow.FindName("CreateButton")
    $pe=$script:mainWindow.FindName("PreviewEditButton"); $pcp=$script:mainWindow.FindName("PreviewCopyPhone")
    $websiteLink=$script:mainWindow.FindName("WebsiteLink")
    $aboutButton=$script:mainWindow.FindName("AboutButton")
    $suggestionsPopup=$script:mainWindow.FindName("SearchSuggestions")
    $suggestionsList=$script:mainWindow.FindName("SuggestionsList")

    # Таймер поиска (debounce 300ms)
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(300)
    $timer.Add_Tick({ $timer.Stop(); Update-TableView })

    # Обработчики событий
    $sb.Add_TextChanged({
        $timer.Stop(); $timer.Start()
        $cb.Visibility = if ($sb.Text.Length -gt 0) { "Visible" } else { "Collapsed" }
        # Плейсхолдер
        if ($null -ne $sp) { $sp.Visibility = if ($sb.Text.Length -gt 0) { "Collapsed" } else { "Visible" } }

        # Автодополнение: показать подсказки
        $query = $sb.Text.Trim().ToLower()
        $suggestionsList.Items.Clear()
        if ($query.Length -ge 2 -and $script:userList.Count -gt 0) {
            $matches = $script:userList | Where-Object {
                ($_.Surname -like "*$query*") -or
                ($_.GivenName -like "*$query*") -or
                ($_.Mail -like "*$query*")
            } | Select-Object -First 8
            foreach ($m in $matches) {
                $display = "$($m.Surname) $($m.GivenName)"
                if ($m.Department -ne '') { $display += " — $($m.Department)" }
                $item = New-Object System.Windows.Controls.ListBoxItem
                $item.Content = $display
                $item.Tag = $m
                [void]$suggestionsList.Items.Add($item)
            }
            if ($suggestionsList.Items.Count -gt 0) {
                $suggestionsPopup.IsOpen = $true
            } else {
                $suggestionsPopup.IsOpen = $false
            }
        } else {
            $suggestionsPopup.IsOpen = $false
        }
    })
    $cb.Add_Click({ $sb.Text=''; $timer.Stop(); $suggestionsPopup.IsOpen=$false; Update-TableView; [void]$sb.Focus() })

    # Выбор подсказки из выпадающего списка
    $suggestionsList.Add_SelectionChanged({
        $selected = $suggestionsList.SelectedItem
        if ($null -ne $selected -and $null -ne $selected.Tag) {
            $emp = $selected.Tag
            $sb.Text = "$($emp.Surname) $($emp.GivenName)"
            $suggestionsPopup.IsOpen = $false
            $timer.Stop(); $timer.Start()
            [void]$sb.Focus()
        }
    })
    $df.Add_SelectionChanged({ Update-TableView })
    $tf.Add_SelectionChanged({ Update-TableView })

    $dg.Add_SelectionChanged({
        $selected = $dg.SelectedItem
        Update-PreviewPanel -Employee $selected

        # Подсветка выделенной строки зелёным цветом
        foreach ($row in $dg.Items) {
            $dgRow = $dg.ItemContainerGenerator.ContainerFromItem($row)
            if ($null -ne $dgRow) {
                if ($row -eq $selected) {
                    $dgRow.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#E8F5D9"))
                } else {
                    $dgRow.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
                }
            }
        }
    })

    $dg.Add_MouseDoubleClick({
        if ($dg.SelectedItem -ne $null -and -not $script:isLoading) {
            Show-EmployeeCard -Employee $dg.SelectedItem -EditMode $false
        }
    })

    # Кнопка редактировать в панели просмотра
    $pe.Add_Click({
        if ($script:selectedEmployee -ne $null -and -not $script:isLoading) {
            if (-not $script:isDomainAdmin) {
                [System.Windows.MessageBox]::Show(
                    "Только доменные администраторы могут редактировать данные.",
                    "Доступ запрещён",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                return
            }
            # После входа — полный доступ к редактированию
            Show-EmployeeCard -Employee $script:selectedEmployee -EditMode $true
        }
    })

    # Кнопка копировать телефон в панели просмотра — отключена (ошибка замыкания)
    $pcp.Visibility = "Collapsed"

    # Ссылка на сайт в шапке
    if ($null -ne $websiteLink) {
        $websiteLink.Add_MouseLeftButtonDown({
            Start-Process "https://zelenoeyabloko.ru"
        })
    }

    # Кнопка "О программе"
    if ($null -ne $aboutButton) {
        $aboutButton.Add_Click({ Show-AboutWindow })
    }

    # Кнопка переключения темы
    $themeBtn=$script:mainWindow.FindName("ThemeToggleButton")
    if ($null -ne $themeBtn) {
        $themeBtn.Add_Click({ Set-Theme -Dark (-not $script:isDarkTheme) })
    }

    $eb.Add_Click({ Export-ToCsv })

    # Кнопка "Создать новый контакт"
    if ($null -ne $crb) {
        $crb.Add_Click({ Show-CreateContactCard })
    }

    # Кнопка email обратной связи
    $feedbackEmail = $script:mainWindow.FindName("FeedbackEmail")
    if ($null -ne $feedbackEmail) {
        $feedbackEmail.Add_MouseLeftButtonDown({
            Start-Process "mailto:it@pepper-group.ru"
        })
    }

    # После входа — пользователь уже является Domain Admin
    # Кнопка «Редактировать» доступна всегда

    # Загрузка логотипа по полному пути
    $logoImg = $script:mainWindow.FindName("LogoImage")
    if ($null -ne $logoImg) {
        $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $exeDir = [System.IO.Path]::GetDirectoryName($exePath)
        $logoPath = [System.IO.Path]::Combine($exeDir, "Logo.png")
        if (Test-Path $logoPath) {
            $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
            $bitmap.BeginInit()
            $bitmap.UriSource = New-Object System.Uri($logoPath, [System.UriKind]::Absolute)
            $bitmap.CacheOption = "OnLoad"
            $bitmap.EndInit()
            $bitmap.Freeze()
            $logoImg.Source = $bitmap
        }
    }

    $script:mainWindow.Add_Loaded({
        $script:refreshTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:refreshTimer.Interval = [TimeSpan]::FromMilliseconds(800)
        $script:refreshTimer.Add_Tick({
            $script:refreshTimer.Stop()
            try {
                Invoke-Refresh
            } catch {
                $si = $script:mainWindow.FindName("StatusIndicator")
                $st = $script:mainWindow.FindName("StatusText")
                if ($null -ne $si) { $si.Fill = "#C0392B" }
                if ($null -ne $st) { $st.Text = "Ошибка загрузки: $($_.Exception.Message)" }
            }
        })
        $script:refreshTimer.Start()
    })

    # Аутентификация при запуске — только Domain Admins
    $loginOk = Show-LoginDialog
    if (-not $loginOk) {
        exit
    }

    $null = $script:mainWindow.ShowDialog()
} catch {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show("Критическая ошибка: $($_.Exception.Message)`n`n$($_.ScriptStackTrace)", "Ошибка запуска",
        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit 1
} finally {
    if ($null -ne $script:refreshTimer -and $script:refreshTimer.IsEnabled) { $script:refreshTimer.Stop() }
    Write-Progress -Activity "Phonebook" -Completed
}






