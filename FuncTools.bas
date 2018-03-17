Attribute VB_Name = "FuncTools"
'*EN*
' This module implements some functional programming solutions
'*RU*
' Модуль реализует некоторые ФП-решения

'*EN*
' Apply array of functions one-by-one to input_data
' Instead of out = f1(f2(f3(f4(f5(x))))) you can create array
' fnames = Array("f1", "f2", "f3", "f4", "f5") (sorry, there are not callbacks)
' and next apply them using for-loop one-by-one to input_data
' This implementation is as close to bash or haskell pipelining as possible
' It's NOT real functional programming, but it's better then Nothing
'*RU*
' Применяет по очереди список функций к входным данным
' Вместо того, чтобы писать out = f1(f2(f3(f4(f5(x))))),
' можно создать массив, содержащий имена функций
' fnames = Array("f1", "f2", "f3", "f4", "f5") (это НЕ коллбэки. В VBA их нет)
' и следом их применить одну за одной ко входным данным
' Эта реализация настолько близка к пайплайнингу bash или Haskell,
' насколько это возможно
' ВНИМАНИЕ: Это НЕ функциональное программирование, но это лучше, чем ничего
Function apply(array_function_names() As Variant, input_data As Variant) As Variant
Dim N As Integer
Dim function_name As Variant
    apply = input_data
    For N = LBound(array_function_names) To UBound(array_function_names)   'reduce
        function_name = array_function_names(N)
        apply = Application.Run(function_name, apply)
    Next N
End Function

'*EN*
' python-like all function all
' If all `conditions` is True, returns True
'*RU*
' Python имеет встроенную функцию all
' Она проходит по всем условиям и возвращает False, как
' только встретит 1й False
Public Function all(conditions() As Variant) As Boolean
Dim i As Integer
    all = True
    For i = LBound(conditions) To UBound(conditions)
        If Not conditions(i) Then
            all = False
            Exit Function
        End If
    Next i
End Function

'*EN*
' python-like all function any
' return True when meet first of `conditions` equals True
'*RU*
' Аналог функции any в Python. Возвращает False только если
' во входном массиве не было обнаружено ни одного True
Public Function anyone(conditions() As Variant) As Boolean
Dim i As Integer
    anyof = False
    For i = LBound(conditions) To UBound(conditions)
        If conditions(i) Then
            anyone = True
            Exit Function
        End If
    Next i
End Function
