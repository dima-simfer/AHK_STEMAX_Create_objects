#SingleInstance force ;�� ���������� � ���, ��� ����� ���� ������� ������ ���� �������.

;������ ��������� ������ ���������� ����� � �������� ��������� � ��������� ��������. ���������, ����� ����� ������� 10-20-100 ����� ��� ������, ������ � �.�.

;�����!!!! ��������� ������, ������� �� ���� ������ ������� ���� � ������ ����� Run with UI Access

;��� 1: ������ � ������������� ���� ��������� ������, ��������� �����������, � ��������������, �������� � �.�.
;��� 2: ������ � �������������� �� ������ ���� ������ ������. �����: �������� ������� ������ ���� ��� �������, ����� ������ �� ���������.
;��� 3: ����� �������� ����� ��������� ������ � ������ ������ ��������������, ����� ������ �������� ������ � ������ ������ � ������� ������.
;��� 4: ��������� ������ (Run with UI Access), ����� �������� ���������������� �����, �������� �������, ��������� ������ ����, �������� ������ ����.


#NoEnv
#SingleInstance force

Gui, +AlwaysOnTop		; ������ ����� ������ ���� ����, ����� �� ���������� �������������� �� �������������� ������.
Gui, font, s8, Verdana  ; Set 8-point Verdana.


;^k::
Gui, Add, Text,, *��� ��� ����������� ��������
Gui, Add, Edit, r1 vObjName w320,
Gui, Add, Text, Section , *��������� ����� �������
Gui, Add, Edit, Number y+1 r1 vStartNum  w160, 
Gui, Add, Text, ys, *�������� ����� �������
Gui, Add, Edit, Number y+1 r1 vEndNum  w150,
Gui, Add, Text, xm, �������� ������� (��������������� � ��������)
Gui, Add, Edit, y+1 r1 vTemplateName w320,
Gui, Add, Checkbox, Section vLoadZones, ��������� ������
Gui, Add, Checkbox, vLoadPersons, ��������� ��������
Gui, Add, Checkbox,ys vLoadInOuts, ��������� �����/������
Gui, Add, Checkbox, vLoadSchedule, ��������� ����������
Gui, Add, GroupBox, xm w320 h1 
Gui, Add, Text, xm, �������� �� �������� �������� `n� ���������� ������� (���)
Gui, Add, Edit, Number x+2 r1 vDelay  w50, 3
Gui, Add, Button, xm+110 Default w100, ����
Gui, Show, Autosize y0 xCenter
return



Button����:
Gui, Submit, NoHide
GuiControlGet, LoadZones, ,, LoadZones
GuiControlGet, LoadPersons, ,, LoadPersons
GuiControlGet, LoadInOuts, ,, LoadInOuts
GuiControlGet, LoadSchedule, ,, LoadSchedule



IfWinExist, STEMAX �������������
		{
			WinActivate
			Sleep 200
		}
		else
		{
			MsgBox 64, ������, ��������� �������-�������������
			return
		}

CurrentNum := StartNum
CB_SELECTSTRING := 0x014d
Template := TemplateName

if ((!ObjName) or (!StartNum) or (!EndNum))
	{
		MsgBox 64, ������, �� ��� ������������ ����(*) ���������
		return
	}	

if (StartNum>=EndNum)
	{
		MsgBox 64, ������, ��������� ����� ������� �� ����� ���� ������ ���� ����� ���������
		return
	}	
	
Loop
{
	
	Click right 35, 135
	Sleep, 300
	Send, {Down}{Enter}
	Sleep, 300

	IfWinExist, �������� �������
		WinActivate

	WinWaitActive, �������� �������,, 5
		if ErrorLevel
		{
    		MsgBox 64, ������, �� ������� ���� �������� �������
			break
		}

	

	SetControlDelay -1

	ControlSetText , Edit1, %ObjName%,,,, 
	Sleep, 50
	
	ControlSetText , Edit2, %CurrentNum% ,,,,
	Sleep, 50

	Control, ChooseString, %Template% , ComboBox1
	;SendMessage, CB_SELECTSTRING,-1, "�������� �������" , ComboBox1	; �������� �������, ������� ����� �������� ������� ��� �������� ��������, ������ �����-������� ������� ������, ���������������� ������ ����.
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

	IfWinExist, ������
		break
	if (CurrentNum>=EndNum)
		break
	CurrentNum++

	Sleep, Delay*1000
	Send, {Home}
}
MsgBox 64, �����, ������ ��������, 2
return