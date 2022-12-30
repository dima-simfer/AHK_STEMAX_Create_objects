#SingleInstance force ;Не спрашивать о том, что может быть запущен только один инстанс.

;Скрипт добавляет нужное количество ячеек с заданным названием и выбранным шаблоном. Актуально, когда нужно создать 10-20-100 ячеек для Болида, Рубежа и т.д.

;ВАЖНО!!!! Запускать скрипт, щёлкнув по нему правой кнопкой мыши и выбрав пункт Run with UI Access

;Шаг 1: создаём в Адмнистраторе одну идеальную ячейку, полностью заполненную, с ответственными, режимами и т.д.
;Шаг 2: Создаём в Администраторе на основе этой ячейки шаблон. ВАЖНО: Название шаблона должно быть без ковычек, иначе шаблон не создастся.
;Шаг 3: Пишем название нашей идеальной ячейки в фильтр сверху Администратора, чтобы убрать ненужные группы и другие ячейки в текущей группе.
;Шаг 4: Запускаем скрипт (Run with UI Access), пишем название свежесоздаваемых ячеек, название шаблона, заполняем другие поля, нажимаем кнопку ПУСК.


#NoEnv
#SingleInstance force

Gui, +AlwaysOnTop		; Окошко будет поверх всех окон, чтобы не переводить Администратора из полноэкранного режима.
Gui, font, s8, Verdana  ; Set 8-point Verdana.


;^k::
Gui, Add, Text,, *Имя для создаваемых объектов
Gui, Add, Edit, r1 vObjName w320,
Gui, Add, Text, Section , *Начальный номер объекта
Gui, Add, Edit, Number y+1 r1 vStartNum  w160, 
Gui, Add, Text, ys, *Конечный номер объекта
Gui, Add, Edit, Number y+1 r1 vEndNum  w150,
Gui, Add, Text, xm, Название шаблона (нечувствительно к регистру)
Gui, Add, Edit, y+1 r1 vTemplateName w320,
Gui, Add, Checkbox, Section vLoadZones, Загружать шлейфы
Gui, Add, Checkbox, vLoadPersons, Загружать Персонал
Gui, Add, Checkbox,ys vLoadInOuts, Загружать входы/выходы
Gui, Add, Checkbox, vLoadSchedule, Загружать расписание
Gui, Add, GroupBox, xm w320 h1 
Gui, Add, Text, xm, Задержка на создание карточки `nи применение шаблона (сек)
Gui, Add, Edit, Number x+2 r1 vDelay  w50, 3
Gui, Add, Button, xm+110 Default w100, Пуск
Gui, Show, Autosize y0 xCenter
return



ButtonПуск:
Gui, Submit, NoHide
GuiControlGet, LoadZones, ,, LoadZones
GuiControlGet, LoadPersons, ,, LoadPersons
GuiControlGet, LoadInOuts, ,, LoadInOuts
GuiControlGet, LoadSchedule, ,, LoadSchedule



IfWinExist, STEMAX Администратор
		{
			WinActivate
			Sleep 200
		}
		else
		{
			MsgBox 64, Ошибка, Запустите Стемакс-Администратор
			return
		}

CurrentNum := StartNum
CB_SELECTSTRING := 0x014d
Template := TemplateName

if ((!ObjName) or (!StartNum) or (!EndNum))
	{
		MsgBox 64, Ошибка, Не все обязательные поля(*) заполнены
		return
	}	

if (StartNum>=EndNum)
	{
		MsgBox 64, Ошибка, Начальный номер объекта не может быть больше либо равен конечному
		return
	}	
	
Loop
{
	
	Click right 35, 135
	Sleep, 300
	Send, {Down}{Enter}
	Sleep, 300

	IfWinExist, Создание объекта
		WinActivate

	WinWaitActive, Создание объекта,, 5
		if ErrorLevel
		{
    		MsgBox 64, Ошибка, Не найдено окно создания объекта
			break
		}

	

	SetControlDelay -1

	ControlSetText , Edit1, %ObjName%,,,, 
	Sleep, 50
	
	ControlSetText , Edit2, %CurrentNum% ,,,,
	Sleep, 50

	Control, ChooseString, %Template% , ComboBox1
	;SendMessage, CB_SELECTSTRING,-1, "Название шаблона" , ComboBox1	; Запасной вариант, указать здесь название шаблона для создания объектов, убрать точку-запятую вначале строки, закомментировать строку выше.
	Sleep, 100
	

	if (LoadZones = 1)
	{
		ControlClick , Button2,,,Left, 1
		Sleep, 50
	}
	
	if (LoadPersons = 1)
	{
		ControlClick , Button3,,,Left, 1
		Sleep, 50
	}
	
	if (LoadInOuts = 1)
	{
		ControlClick , Button4,,,Left, 1
		Sleep, 50
	}
	
	if (LoadSchedule = 1)
	{
		ControlClick , Button5,,,Left, 1
		Sleep, 50
	}
	
	ControlClick, Button6

	IfWinExist, Ошибка
		break
	if (CurrentNum>=EndNum)
		break
	CurrentNum++

	Sleep, Delay*1000
	Send, {Home}
}
MsgBox 64, Конец, Скрипт выполнен, 2
return