Dim objFSO

'http://www.rz.uni-kiel.de/pc/office-kms/
'http://www.rz.uni-kiel.de/pc/office2013-kms/

'KMS Client Keys http://www.infrastrukturhelden.de/microsoft-infrastruktur/key-management-service-kms-client-seriennummern-updated.html

' find KMS Server via DNS SRV Record
autodetect="True"
' or connect to this specified KMS Server
kms_server="kms.mydomain"
kms_port="1688"

 
Set oShell = CreateObject("WScript.Shell")

'Get Office Installpath
OfficeVersion = oShell.RegRead ("HKLM\SOFTWARE\Microsoft\Office\Common\LastAccessInstall")
If OfficeVersion < 14 Then
	Msgbox "Abbruch: KMS-Lizenzierung funktioniert erst ab Office 2010"
	Wscript.Quit
End If

ProductName = oShell.RegRead ("HKLM\SOFTWARE\Microsoft\Office\" + Trim(OfficeVersion) + ".0\Registration\{6F327760-8C5C-417C-9B61-836A98287E0C}\ProductName")
OfficePath = oShell.RegRead ("HKLM\SOFTWARE\Microsoft\Office\" + Trim(OfficeVersion) + ".0\Common\InstallRoot\Path")
KMSScript = OfficePath + "OSPP.VBS"


Select Case ProductName
' Office 2010 ----------------------------------------
	Case "Microsoft Office Professional Plus 2010"
		kms_serial = "VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB"
	Case "Microsoft Office Standard 2010"
		kms_serial = "V7QKV-4XVVR-XYV4D-F7DFM-8R6BM"

' Office 2013 ----------------------------------------
	Case "Microsoft Office 2013 Professional Plus"
		kms_serial = "PGD67-JN23K-JGVWV-KTHP4-GXR9G"
	Case "Microsoft Office 2013 Standard"
		kms_serial = "KBKQT-2NMXY-JJWGP-M62JB-92CD4"

' Office 2016 ----------------------------------------
	Case "Microsoft Office 2016 Professional Plus"
		kms_serial = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
	Case "Microsoft Office 2016 Standard"
		kms_serial = "JNRGM-WHDWX-FJJG3-K47QV-DRTFM"

' Office 2019 ----------------------------------------
	Case "Microsoft Office 2019 Professional Plus"
		kms_serial = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
	Case "Microsoft Office 2019 Standard"
		kms_serial = "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK"

' Sharepoint Workspace 2010 ----------------------------------------
    Case "Microsoft SharePoint Workspace 2010"
		kms_serial = "QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4"

' Office 2010 Products ----------------------------------------
    Case "Microsoft Access 2010"
		kms_serial = "V7Y44-9T38C-R2VJK-666HK-T7DDX"
    Case "Microsoft Excel 2010"
		kms_serial = "H62QG-HXVKF-PP4HP-66KMR-CW9BM"
    Case "Microsoft InfoPath 2010"
		kms_serial = "K96W8-67RPQ-62T9Y-J8FQJ-BT37T"
    Case "Microsoft OneNote 2010"
		kms_serial = "Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX"
    Case "Microsoft Outlook 2010"
		kms_serial = "7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ"
    Case "Microsoft PowerPoint 2010"
		kms_serial = "RC8FX-88JRY-3PF7C-X8P67-P4VTT"
    Case "Microsoft Project Professional 2010"
		kms_serial = "YGX6F-PGV49-PGW3J-9BTGG-VHKC6"
    Case "Microsoft Project Standard 2010"
		kms_serial = "4HP3K-88W3F-W2K3D-6677X-F9PGB"
    Case "Microsoft Publisher 2010"
		kms_serial = "BFK7F-9MYHM-V68C7-DRQ66-83YTP"
    Case "Microsoft Word 2010"
		kms_serial = "HVHB3-C6FV7-KQX9W-YQG79-CRY7T"
    Case "Microsoft Visio Premium 2010"
		kms_serial = "D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ"
    Case "Microsoft Visio Professional 2010"
		kms_serial = "7MCW8-VRQVK-G677T-PDJCM-Q8TCP"
    Case "Microsoft Visio Standard 2010"
		kms_serial = "767HD-QGMWX-8QTDB-9G3R2-KHFGJ"

' Office 2013 Products ----------------------------------------
    Case "Microsoft Project 2013 Professional"
		kms_serial = "FN8TT-7WMH6-2D4X9-M337T-2342K"
    Case "Microsoft Project 2013 Standard"
		kms_serial = "6NTH3-CW976-3G3Y2-JK3TX-8QHTT"
    Case "Microsoft Visio 2013 Professaional"
		kms_serial = "C2FG9-N6J68-H8BTJ-BW3QX-RM3B3"
    Case "Microsoft Visio 2013 Standard"
		kms_serial = "J484Y-4NKBF-W2HMG-DBMJC-PGWR7"
    Case "Microsoft Access 2013"
		kms_serial = "NG2JY-H4JBT-HQXYP-78QH9-4JM2D"
    Case "Microsoft Excel 2013"
		kms_serial = "VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB"
    Case "Microsoft InfoPath 2013"
		kms_serial = "DKT8B-N7VXH-D963P-Q4PHY-F8894"
    Case "Microsoft Lync 2013"
		kms_serial = "2MG3G-3BNTT-3MFW9-KDQW3-TCK7R"
    Case "Microsoft OneNote 2013"
		kms_serial = "TGN6P-8MMBC-37P2F-XHXXK-P34VW"
    Case "Microsoft Outlook 2013"
		kms_serial = "QPN8Q-BJBTJ-334K3-93TGY-2PMBT"
    Case "Microsoft PowerPoint 2013"
		kms_serial = "4NT99-8RJFH-Q2VDH-KYG2C-4RD4F"
    Case "Microsoft Publisher 2013"
		kms_serial = "PN2WF-29XG2-T9HJ7-JQPJR-FCXK4"
    Case "Microsoft Word 2013"
		kms_serial = "6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7"

' Office 2016 Products ----------------------------------------
    Case "Microsoft Project 2016 Professional"
		kms_serial = "YG9NW-3K39V-2T3HJ-93F3Q-G83KT"
    Case "Microsoft Project 2016 Standard"
		kms_serial = "GNFHQ-F6YQM-KQDGJ-327XX-KQBVC"
    Case "Microsoft Visio 2016 Professaional"
		kms_serial = "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK"
    Case "Microsoft Visio 2016 Standard"
		kms_serial = "7WHWN-4T7MP-G96JF-G33KR-W8GF4"
    Case "Microsoft Access 2016"
		kms_serial = "GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW"
    Case "Microsoft Excel 2016"
		kms_serial = "9C2PK-NWTVB-JMPW8-BFT28-7FTBF"
    Case "Microsoft OneNote 2016"
		kms_serial = "DR92N-9HTF2-97XKM-XW2WJ-XW3J6"
    Case "Microsoft Outlook 2016"
		kms_serial = "R69KK-NTPKF-7M3Q4-QYBHW-6MT9B"
    Case "Microsoft PowerPoint 2016"
		kms_serial = "J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6"
    Case "Microsoft Publisher 2016"
		kms_serial = "F47MM-N3XJP-TQXJ9-BP99D-8K837"
    Case "Microsoft Skype for Business 2016"
		kms_serial = "869NQ-FJ69K-466HW-QYCP2-DDBV6"
    Case "Microsoft Word 2016"
		kms_serial = "WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6"

' Office 2019 Products ----------------------------------------
    Case "Microsoft Project 2019 Professional"
		kms_serial = "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B"
    Case "Microsoft Project 2019 Standard"
		kms_serial = "C4F7P-NCP8C-6CQPT-MQHV9-JXD2M"
    Case "Microsoft Visio 2019 Professaional"
		kms_serial = "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB"
    Case "Microsoft Visio 2019 Standard"
		kms_serial = "7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2"
    Case "Microsoft Access 2019"
		kms_serial = "9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT"
    Case "Microsoft Excel 2019"
		kms_serial = "TMJWT-YYNMB-3BKTF-644FC-RVXBD"
    Case "Microsoft Outlook 2019"
		kms_serial = "7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK"
    Case "Microsoft PowerPoint 2019"
		kms_serial = "RRNCX-C64HY-W2MM7-MCH9G-TJHMQ"
    Case "Microsoft Publisher 2019"
		kms_serial = "G2KWX-3NW6P-PY93R-JXK2T-C9Y9V"
    Case "Microsoft Skype for Business 2019"
		kms_serial = "NCJ33-JHBBY-HTK98-MYCV8-HMKHJ"
    Case "Microsoft Word 2019"
		kms_serial = "PBX3G-NWMT6-Q7XBW-PYJGG-WXD33"

    Case Else
		Msgbox ("Abbruch: kein unterstütztes Office Produkt gefunden! ->" + ProductName)
		WScript.Quit
End Select


Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(OfficePath) Then
	If (objFSO.FileExists(KMSScript)) Then
		if autodetect = "True" Then
			Msgbox ("setze KMS Server auf Autodetect ...")
			oShell.run "cscript " + KMSScript + " /remhst",1,True
		else
			Msgbox ("schreibe Adresse des KMS Server in die Registry ...")
			oShell.run "cscript " + KMSScript + " /sethst:" + kms_server,1,True
			oShell.run "cscript " + KMSScript + " /setprt:" + kms_port,1,True
		End if
		oShell.run "cscript " + KMSScript + " /inpkey:" + kms_serial,1,True
		oShell.run "cscript " + KMSScript + " /act",1,True
		Msgbox "Erledigt :-)"
	else
		Msgbox ("Abbruch: KMS-Client Script nicht gefunden")
		WScript.Quit
	End If
Else
	MsgBox ("Abbruch: Konnte Office Pfad nicht finden!")
	WScript.Quit
End If

'Set wshshell = WScript.CreateObject ("wscript.shell")
' wshshell.run "c:\Windows\system32\cmd.exe", 6, True
' set wshshell = nothing 
'
' der Parameter 6: Minimiert das Fenster
' 0: versteckt das Fenster und aktiviert ein anderes
' 1: aktiviert und zeigt ein Fenster
' 2: aktiviert und minimiert das Fenster
' 3: aktiviert und maximiert das Fenster
' 4: zeigt das Fenster in seiner letzen Position, das aktive Fenster bleibt aktiv
' 5: zeigt das Fenster in seiner letzen grösse und Position
' 6: minimiert das Fenster und aktiviert ein anderes
' 7: minimiert das Fenster, das aktive Fenster bleibt aktiv
' 8: zeigt das Fenster in seiner letzen Position, das aktive Fenster bleibt aktiv
' 9: stellt ein minimiertes Fenster wieder in seinen ursprünglichen Zustand
' 10: setzt das Fenster gleich dem Programm 
'
' True: Script wartet, bis der Task beendet wird, False: Script läuft weiter

