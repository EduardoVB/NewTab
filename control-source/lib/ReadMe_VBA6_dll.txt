Tab is a reserved keyword in VB6, but you can remove that restriction.
To be able to compile with Tab property, you need to replace VBA6.DLL with this version.
VBA6.DLL is in VB6's installation folder, usually:
C:\Program Files (x86)\Microsoft Visual Studio\VB98\

Note: this DLL is for SP6, so it is assumed that you already have SP6 (Service Pack 6) installed for your VB6 IDE.

Also, you need to change the line:
#Const COMPILE_WITH_TAB_PROPERTY = 0
to:
#Const COMPILE_WITH_TAB_PROPERTY = 1

That line is almost at the end of the NewTab UserControl code module.

In addition, if you want to preserve the binary compatibility, you'll need to go the the Project menu, Properties, Component tab, and set the "Binary compatibility" option.