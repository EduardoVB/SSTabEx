For getting help on the control, read the file named "tabexctl_reference". The same text is available in a property page of the control.

If you are going to add this control to an standard EXE project, these are the files needed:

ctl\uscSSTabEx.ctl 			[main file]
ctl\uscSSTabEx.ctx 			[main file]
subclass\cIBSSubclass.cls		[for subclassing]
subclass\mBSPropsDB.bas			[for subclassing]
subclass\mBSSubclass.bas		[for subclassing]
misc\cDlg.cls				[used in a property page]
misc\frmSSTabExSelectControl.frm	[used in a property page]
ptp\ptpSSTabExGeneral.pag		[property page]
ptp\ptpSSTabExGeneral.pgx		[property page]
ptp\ptpSSTabExHelp.pag			[property page]
ptp\ptpSSTabExHelp.pgx			[property page]
ptp\ptpSSTabExTabs.pag			[property page]
ptp\ptpSSTabExTabs.pgx			[property page]

In an EXE project, only the first 5 files are strictly necessary, but if you don't add the rest you'll lose the property pages (and their design time helper functions, like changing controls from one tab to another).

The compiled OCX is in the folder control-bin.

If you are going to compile your own OCX, please change the project name and the OCX name to avoid conflics with the "official" OCX or with other OCX that other developers generated (known as DLL hell).

Better if you use Side-by-side assemblies (SxS) to avoid dll hell (conflicting versions), even using the "official" OCX versions.

For the SxS manifest file, the XML text for the OCX that is in the control-bin folder is:

  <file name="Bin\TabExC01.ocx">
    <typelib tlbid="{EA478B61-D9EC-47F6-BB21-95A533AF2251}" version="1.0" flags="control" helpdir="" />
    <comClass clsid="{A1462394-2F1F-4B72-AABF-31DC289A86AE}" tlbid="{EA478B61-D9EC-47F6-BB21-95A533AF2251}" threadingModel="Apartment" progid="TabExCtl.SSTabEx" miscStatusIcon="recomposeonresize,cantlinkinside,insideout,activatewhenvisible,simpleframe,setclientsitefirst" description="" />
  </file>

