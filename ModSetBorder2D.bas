Attribute VB_Name = "ModSetBorder2D"
Option Explicit

'SetBorder2D�E�E�E���ꏊ�FFukamiAddins3.ModBorder

'�錾�Z�N�V����������������������������������������������������������
'-----------------------------------
'���ꏊ:LineStyleEnum
Public Enum LineStyleEnum '���̃X�^�C��
    ���� = xlContinuous
    �Ȃ� = xlNone
    �j�� = xlDash
    ��_���� = xlDashDot
    ��_���� = xlDashDotDot
    �_�� = xlDot
    ��d�� = xlDouble
    �Δj�� = xlSlantDashDot
End Enum
'-----------------------------------
'���ꏊ:LineWeightEnum
Public Enum LineWeightEnum '���̑���
    ���� = xlThick
    ������ = xlMedium
    �א� = xlThin
    �ɍא� = xlHairline
End Enum
'�錾�Z�N�V�����I��������������������������������������������������������

Public Sub SetBorder2D(TargetCell As Range, BaseCol As Long, _
                Optional EdgeLineStyle As LineStyleEnum = LineStyleEnum.����, _
                Optional EdgeLineWeight As LineWeightEnum = LineWeightEnum.������, _
                Optional InsideHorizontalLineStyle As LineStyleEnum = LineStyleEnum.�_��, _
                Optional InsideHorizontalLineWeight As LineWeightEnum = LineWeightEnum.�א�, _
                Optional InsideVerticalLineStyle As LineStyleEnum = LineStyleEnum.����, _
                Optional InsideVerticalLineWeight As LineWeightEnum = LineWeightEnum.�א�)

'�w��Z���͈͂����₷���悤�Ɍr����ݒ肷��
'���ɂ����āA�l���؂�ւ��Ƃ��낾���������𑾂������肷��B
'20210917

'����
'TargetCell                  �E�E�E�ΏۂƂ���Z���͈̔�(Range�^)
'BaseCol                     �E�E�E��̗�i�Ώ۔͈̓Z���̍����牽�Ԗڂ��j(Long�^)
'[EdgeLineStyle]             �E�E�E�O���r���̃X�^�C��    �i�f�t�H���g�͎����j
'[EdgeLineWeight]            �E�E�E�O���r���̑���        �i�f�t�H���g�͒������j
'[InsideHorizontalLineStyle] �E�E�E���������r���̃X�^�C���i�f�t�H���g�͓_���j
'[InsideHorizontalLineWeight]�E�E�E���������r���̑���    �i�f�t�H���g�͍א��j
'[InsideVerticalLineStyle]   �E�E�E���������r���̃X�^�C���i�f�t�H���g�͎����j
'[InsideVerticalLineWeight]  �E�E�E���������r���̑���    �i�f�t�H���g�͍א��j
    
    Dim BaseList
    Dim I          As Long
    Dim N          As Long
    Dim InputSheet As Worksheet
    Dim Rs         As Long
    Dim Re         As Long
    Dim Cs         As Long
    Dim Ce         As Long
    Set InputSheet = TargetCell.Parent '�Ώۂ̃V�[�g�擾
    Rs = TargetCell(1).Row
    Cs = TargetCell(1).Column
    Re = TargetCell(TargetCell.Count).Row
    Ce = TargetCell(TargetCell.Count).Column
    
    If Rs = Re Then '�Z���͈͂�1�s�������Ȃ��ꍇ�͒��ڐݒ�
        With TargetCell
            .Borders.LineStyle = EdgeLineStyle
            .Borders.Weight = EdgeLineWeight
            .Borders(xlInsideHorizontal).LineStyle = InsideHorizontalLineStyle
            .Borders(xlInsideHorizontal).Weight = InsideHorizontalLineWeight
            .Borders(xlInsideVertical).LineStyle = InsideVerticalLineStyle
            .Borders(xlInsideVertical).Weight = InsideVerticalLineWeight
        End With
        Exit Sub
    End If
    
    '��̈ꎟ���z����擾
    With InputSheet
        BaseList = .Range(.Cells(Rs, Cs + BaseCol - 1), .Cells(Re, Cs + BaseCol - 1)).Value
    End With
    BaseList = Application.Transpose(BaseList)
    
    N = UBound(BaseList, 1)
    Dim StartCell As Range
    Dim EndCell   As Range
    Dim Hantei    As Boolean
    
    Application.ScreenUpdating = False '��ʍX�V�����ō�����
    For I = 1 To N
        If I = 1 Then
            Set StartCell = InputSheet.Cells(Rs, Cs)
        End If
        
        If I = N Then
            Set EndCell = InputSheet.Cells(Re, Ce)
            Hantei = True
        ElseIf BaseList(I) <> BaseList(I + 1) Then
            Set EndCell = InputSheet.Cells(Rs + I - 1, Ce)
            Hantei = True
        Else
            Hantei = False
        End If
    
        If Hantei = True Then
            With Range(StartCell, EndCell)
                .Borders.LineStyle = EdgeLineStyle
                .Borders.Weight = EdgeLineWeight
                .Borders(xlInsideHorizontal).LineStyle = InsideHorizontalLineStyle
                .Borders(xlInsideHorizontal).Weight = InsideHorizontalLineWeight
                .Borders(xlInsideVertical).LineStyle = InsideVerticalLineStyle
                .Borders(xlInsideVertical).Weight = InsideVerticalLineWeight
            End With
            Set StartCell = InputSheet.Cells(Rs + I, Cs)
        End If
    Next I
    Application.ScreenUpdating = True '��ʍX�V�����̉���
    
End Sub


