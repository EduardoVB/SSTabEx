For getting help on the control, read the file that is in the folder "others" and is named "Help text.txt". The same text is available in a property page of the control.

If you are going to add this control to an standard EXE project, these are the needed files:

usc / uscSSTabEx.ctl
usc / uscSSTabEx.ctx
Subclass / isubclass.cls
Subclass / mPropsDB.bas
Subclass / mSubclass.bas
Subclass / subclass.cls
others / cCDlg.cls
ptp / ptpSSTabExGeneral.pag
ptp / ptpSSTabExGeneral.pgx
ptp / ptpSSTabExHelp.pag
ptp / ptpSSTabExHelp.pgx
ptp / ptpSSTabExTabs.pag
ptp / ptpSSTabExTabs.pgx

Since this control is subclassed, for better IDE stability while in design mode, uncomment the line:
'#Const NOSUBCLASSINIDE = True
that line is in the SSTabEx control code module.
Doing so, you'll lose some features in the IDE, being the most important one the ability to change the selected tab by clicking while in design mode.
You can still change the selected tab by changing the TabSel property setting from the property window.

If you are going to use this control in a component (like it is released), if you run it uncompiled you have the same situation.
That's why it's better to compile the OCX file and remove the project SSTabRpl from the project group.

To compile the OCX file:
Got to File/Make... and make the OCX.
Save the project.
Close the IDE.
Copy the *.OCX that you generated to the folder cmp
Rename it as *.cmp
Open the project in the IDE.
Go to Project/Properties, and in the Components tab, in Binary compatibility select the file *.cmp in the folder cmp.
Click OK.
Save the project.

Doing so, if you make changes to the code and re-compile, the programs using the OCX will be automatically updated when the are loaded in the iDE.

If the changes that you made broke the compatibility (for example changing a property name), when trying to compile VB6 will complain.
In that case, cancel (the compilation) and go to Project/Properties, and in the Components tab and select "Project compatibility. Click OK.
Generate the OCX.
Close the IDE.
Copy the *.OCX that you generated to the folder cmp
Rename it as *.cmp replacing the old one.
Open the project in the IDE.
Go to Project/Properties, and in the Components tab, in Binary compatibility select the file *.cmp in the folder cmp.
Click OK.
Save the project.

You'll need to do the steps detailed above every time that you break the binary compatibility.
Doing so, all the projects that are using the OCX will be updated to the new version.
Of course if you, for example, changed the name of a property, you'll have to also change the name of that property in the code of the EXE's, if the property is referenced in the code.
