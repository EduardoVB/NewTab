Steps for replacing SSTab with NewTab in a project:

1 - Make a backup of your original project.
2 - Open the project in the IDE and from menu Project, Components, select NewTab for VB6, then click OK.
3 - Save the project and close the IDE.
4 - Not required but if you open the vbp file with Notepad now you should see this line:
Object={66E63055-5A66-4C79-9327-4BC077858695}#14.0#0; NewTab01.ocx
5 - Now you need to open with Notepad or other text editor each frm file and replace the text 'TabDlg.SSTab' with 'NewTabCtl.NewTab'
6 - In Notepad, use Save As and change the Encoding from UTF-8 to ANSI.
7 - Alternatively to steps 5 and 6, if you have many forms, you can use a program to make a global text replace in all the *.frm files in a folder. You can use TextRep available here:  https://no-nonsense-software.com/freeware
8 - Open the project in the IDE and go to menu Project, Components, and deselect Microsoft Tabbed Dialog Control 6.0, click OK.
9 - Save the project.


Code for the manifest file to make a SxS installation:

  <file name="NewTab01.ocx">
    <typelib tlbid="{66E63055-5A66-4C79-9327-4BC077858695}" version="14.0" flags="control" helpdir="" />
    <comClass clsid="{D9DACE39-8348-4FC3-8BBF-9178D817C34B}" tlbid="{66E63055-5A66-4C79-9327-4BC077858695}" threadingModel="Apartment" progid="NewTabCtl.NewTab" miscStatusIcon="recomposeonresize,cantlinkinside,insideout,activatewhenvisible,alignable,simpleframe,setclientsitefirst" description="NewTab: tabbed control" />
  </file>
