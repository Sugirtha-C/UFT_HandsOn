varB=Browser("OrangeHRM").Page("OrangeHRM").WebEdit("username").GetToProperty("name")
msgBox varB


Set varC=Browser("OrangeHRM").Page("OrangeHRM").WebEdit("username")
Set props=varC.GetTOProperties
propsCount=props.Count
print propsCount

For i = 0 To propsCount-1 
propName=props(i).Name 'Name=Returns the name of the property 
propValue=props(i).Value 'Value= retusn the value of property
MsgBox propName &"="&propValue
	
Next




