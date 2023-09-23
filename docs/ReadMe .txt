Steps for replacing SSTab with NewTab in a project:

1 - Make a backup of your original project.
2 - Open the project in the IDE and from menu Project, Components, select NewTab for VB6, then click OK.
3 - Save the project and close the IDE.
4 - Not required but if you open the vbp file with Notepad now you should see this line:
Object={66E63055-5A66-4C79-9327-4BC077858695}#3.0#0; NewTab01.ocx
5 - Now you need to open with Notepad or other text editor each frm file and replace the text 'TabDlg.SSTab' with 'NewTabCtl.NewTab'
If you have many forms you can use a program to make a global text replace in all the *.frm files in a folder, I use TextRep available here:  https://no-nonsense-software.com/freeware
6 - Open the project in the IDE and go to menu Project, Components, and deselect Microsoft Tabbed Dialog Control 6.0, click OK.
7 - Save the project.


Code for the manifest file to make a SxS installation:

  <file name="NewTab01.ocx">
    <typelib tlbid="{66E63055-5A66-4C79-9327-4BC077858695}" version="3.1" flags="control" helpdir="" />
    <comClass clsid="{6C9D299A-B6CC-40B8-A155-4D3F154F6584}" tlbid="{66E63055-5A66-4C79-9327-4BC077858695}" threadingModel="Apartment" progid="NewTabCtl.NewTab" miscStatusIcon="recomposeonresize,cantlinkinside,insideout,activatewhenvisible,simpleframe,setclientsitefirst" description="Tabbed control for VB6." />
  </file>
