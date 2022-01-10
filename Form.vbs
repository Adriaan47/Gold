Private Sub UserForm_Initialize()

'Empty NameTextBox
NameTextBox.Value = ""

'Empty PhoneTextBox
PhoneTextBox.Value = ""

'Empty CityListBox
CityListBox.Clear

'Fill CityListBox
With CityListBox
    .AddItem "San Francisco"
    .AddItem "Oakland"
    .AddItem "Richmond"
End With

'Empty DinnerComboBox
DinnerComboBox.Clear

'Fill DinnerComboBox
With DinnerComboBox
    .AddItem "Italian"
    .AddItem "Chinese"
    .AddItem "Frites and Meat"
End With

'Uncheck DataCheckBoxes
DateCheckBox1.Value = False
DateCheckBox2.Value = False
DateCheckBox3.Value = False

'Set no car as default
CarOptionButton2.Value = True

'Empty MoneyTextBox
MoneyTextBox.Value = ""

'Set Focus on NameTextBox
NameTextBox.SetFocus

End Sub