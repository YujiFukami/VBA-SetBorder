Attribute VB_Name = "ModSetBorder2D"
Option Explicit

'SetBorder2D・・・元場所：FukamiAddins3.ModBorder

'宣言セクション※※※※※※※※※※※※※※※※※※※※※※※※※※※
'-----------------------------------
'元場所:LineStyleEnum
Public Enum LineStyleEnum '線のスタイル
    実線 = xlContinuous
    なし = xlNone
    破線 = xlDash
    一点鎖線 = xlDashDot
    二点鎖線 = xlDashDotDot
    点線 = xlDot
    二重線 = xlDouble
    斜破線 = xlSlantDashDot
End Enum
'-----------------------------------
'元場所:LineWeightEnum
Public Enum LineWeightEnum '線の太さ
    太線 = xlThick
    中太線 = xlMedium
    細線 = xlThin
    極細線 = xlHairline
End Enum
'宣言セクション終了※※※※※※※※※※※※※※※※※※※※※※※※※※※

Public Sub SetBorder2D(TargetCell As Range, BaseCol As Long, _
                Optional EdgeLineStyle As LineStyleEnum = LineStyleEnum.実線, _
                Optional EdgeLineWeight As LineWeightEnum = LineWeightEnum.中太線, _
                Optional InsideHorizontalLineStyle As LineStyleEnum = LineStyleEnum.点線, _
                Optional InsideHorizontalLineWeight As LineWeightEnum = LineWeightEnum.細線, _
                Optional InsideVerticalLineStyle As LineStyleEnum = LineStyleEnum.実線, _
                Optional InsideVerticalLineWeight As LineWeightEnum = LineWeightEnum.細線)

'指定セル範囲を見やすいように罫線を設定する
'基準列において、値が切り替わるところだけ水平線を太くしたりする。
'20210917

'引数
'TargetCell                  ・・・対象とするセルの範囲(Range型)
'BaseCol                     ・・・基準の列（対象範囲セルの左から何番目か）(Long型)
'[EdgeLineStyle]             ・・・外側罫線のスタイル    （デフォルトは実線）
'[EdgeLineWeight]            ・・・外側罫線の太さ        （デフォルトは中太線）
'[InsideHorizontalLineStyle] ・・・内側水平罫線のスタイル（デフォルトは点線）
'[InsideHorizontalLineWeight]・・・内側水平罫線の太さ    （デフォルトは細線）
'[InsideVerticalLineStyle]   ・・・内側垂直罫線のスタイル（デフォルトは実線）
'[InsideVerticalLineWeight]  ・・・内側垂直罫線の太さ    （デフォルトは細線）
    
    Dim BaseList
    Dim I          As Long
    Dim N          As Long
    Dim InputSheet As Worksheet
    Dim Rs         As Long
    Dim Re         As Long
    Dim Cs         As Long
    Dim Ce         As Long
    Set InputSheet = TargetCell.Parent '対象のシート取得
    Rs = TargetCell(1).Row
    Cs = TargetCell(1).Column
    Re = TargetCell(TargetCell.Count).Row
    Ce = TargetCell(TargetCell.Count).Column
    
    If Rs = Re Then 'セル範囲が1行分しかない場合は直接設定
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
    
    '基準の一次元配列を取得
    With InputSheet
        BaseList = .Range(.Cells(Rs, Cs + BaseCol - 1), .Cells(Re, Cs + BaseCol - 1)).Value
    End With
    BaseList = Application.Transpose(BaseList)
    
    N = UBound(BaseList, 1)
    Dim StartCell As Range
    Dim EndCell   As Range
    Dim Hantei    As Boolean
    
    Application.ScreenUpdating = False '画面更新解除で高速化
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
    Application.ScreenUpdating = True '画面更新解除の解除
    
End Sub


