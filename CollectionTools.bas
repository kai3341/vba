Attribute VB_Name = "CollectionTools"
Option Explicit

'Выполняет проверку, существует элемент в коллекции или нет
'Имеет смысл для странных коллекций, не имеющих метода Exists
Public Function hasItem(Target As Variant, key As Variant) As Boolean
hasItem = False
On Error GoTo ExitFunction
Call Target.Item(key)
hasItem = True
ExitFunction:
End Function
