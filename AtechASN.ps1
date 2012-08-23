#########################################################
#ASN Generation Script for Atech						#
#Written by Anthony Sinatra								#
#Written for PerTronix									#
#Build 0.2.0 - Beta										#
#########################################################

#########################################################
#Check and Build required directories					#
#########################################################

if ((Test-Path $home\Documents\ASN) -eq $false)
{
    md $home\Documents\ASN
}

#########################################################
#Set Required Paths										#
#														#
#														#
#Setting the Continue Variable. If this isn't set code	#
#will exit!												#
#########################################################

$Continue = "Yes"
$rootpath = "$home\Documents\iMacros\Macros"
$SaveDir = "\\snap\share1\ASN\Atech"
$CSVDir = "\\snap\share1\ASN\Atech\CSV"
$DownloadDir = "$home\Downloads"
$Date = get-date -uformat %m%d%Y

#########################################################
#Remove Old files from older versions					#
#########################################################

if ((Test-Path $SaveDir\$Date.csv) -eq $true)
{
    Remove-Item $SaveDir\$Date.csv
}
if ((Test-Path $rootpath\AtechASN.iim) -eq $true)
{
    Remove-Item $rootpath\AtechASN.iim
}

#########################################################
#Begin Code!											#
#########################################################

while($Continue -eq "Yes")
{
	#####################################################
	#region functions									#
	#####################################################
	
	#####################################################
	# load WinForms										#
	#####################################################
	
	[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	
	#####################################################
	# create form										#
	#####################################################
	
	$form = New-Object Windows.Forms.Form
	$form.text = "Choose Division"
	$form.top = 10
	$form.left = 10
	$form.height = 250
	$form.width = 350

	#####################################################
	# create label										#
	#####################################################
	
	$label = New-Object Windows.Forms.Label
	$label.text = "Which Division would you like to create ASNs for?"
	$label.height = 275
	$label.width = 100
	$label.top = 2
	$label.left = 25
	$form.controls.add($label)
	
	#####################################################
	# create button										#
	#####################################################
	
	$button = New-Object Windows.Forms.Button
	$button.text = "Run!"
	$button.height = 40
	$button.width = 70
	$button.top = 150
	$button.left = 150
	$form.controls.add($button)
	
	#####################################################
	# create radiobutton								#
	#####################################################
	
	$RadioButton = New-Object Windows.Forms.radiobutton
	$RadioButton.text = "Ignition"
	$RadioButton.height = 20
	$RadioButton.width = 150
	$RadioButton.top = 2
	$RadioButton.left = 150
	$form.controls.add($RadioButton)

	#####################################################
	# create radiobutton1								#
	#####################################################
	
	$radiobutton1 = New-Object Windows.Forms.radiobutton
	$RadioButton1.text = "Exhaust"
	$RadioButton1.height = 20
	$RadioButton1.width = 150
	$RadioButton1.top = 30
	$RadioButton1.left =150
	$form.controls.add($RadioButton1)

	#####################################################
	# create radiobutton2								#
	#####################################################
	
	$radiobutton2 = New-Object Windows.Forms.radiobutton
	$RadioButton2.text = "Private Label"
	$RadioButton2.height = 20
	$RadioButton2.width = 150
	$RadioButton2.top = 58
	$RadioButton2.left =150
	$form.controls.add($RadioButton2)

	#####################################################
	# create event handler for button					#
	#####################################################
	
	$event = {
		if($radiobutton.checked){$Division = "Ignition"}
		if($radiobutton1.checked){$Division = "Exhaust"}
		if($radiobutton2.checked){$Division = "PrivateLabel"}
		$form.Close()
	}

	#####################################################
	# attach event handler								#
	#####################################################
	
	$button.Add_Click($event)

	#####################################################
	# attach controls to form							#
	#####################################################
	
	$form.controls.add($button)
	$form.controls.add($label)
	$form.controls.add($textbox)

	$form.showdialog()



	#####################################################
	#Start Ignition block of code						#
	#####################################################
	
	if ($Division -eq "Ignition")
	{
		if ((Test-Path $DownloadDir\AtechIgnitionASN.csv) -eq $true)
		{
		
			#############################################
			#Remove old files							#
			#############################################
			
			if ((Test-Path $rootpath\AtechIgnitionASN.iim) -eq $true)
			{    
				Remove-Item $rootpath\AtechIgnitionASN.iim
			}
			if ((Test-Path $rootpath\AtechIgnitionLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechIgnitionLogin.iim
			}
			
			#############################################
			#Login information							#
			#############################################
			
			$username = Read-Host 'Please input your Username for Ignition'
			$pass = Read-Host 'Please input your password?' -AsSecureString
			
			#############################################
			#Convert password to plain text variable	#
			#############################################
			
			$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))
			
			#############################################
			#Build Login Macro 							#
			#############################################
			
			Add-Content $rootpath\AtechIgnitionLogin.iim "TAB T=1 `nURL GOTO=https://b2b.atechmotorsports.com/default.asp"
			Add-Content $rootpath\AtechIgnitionLogin.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmLogin ATTR=NAME:UserName CONTENT=$username"
			Add-Content $rootpath\AtechIgnitionLogin.iim "SET !ENCRYPTION NO"
			Add-Content $rootpath\AtechIgnitionLogin.iim "TAG POS=1 TYPE=INPUT:PASSWORD FORM=NAME:frmLogin ATTR=NAME:Password CONTENT=$password"
			Add-Content $rootpath\AtechIgnitionLogin.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmLogin ATTR=VALUE:Go"
			$password = $null
			
			
			#############################################
			#Run Login Macro							#
			#############################################
			
			if ((Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe") -eq $true)
			{
				$cmdLine = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
			}else
			{
				$cmdLine = "C:\Program Files\Mozilla Firefox\firefox.exe"
			}
			$args = "imacros://run/?m=AtechIgnitionLogin.iim"
			start-process $cmdLine $args
			Get-Process | ? {$_.Name -like "firefox"} | %{$_.Close()}
    
			#############################################
			#Ready to move to the Automation?			#
			#############################################
			
#			$a = new-object -comobject wscript.shell
#			$b = $a.popup("Did the script log on to Atech?",0,"Ignition Login Box",1)
			$b = [System.Windows.Forms.MessageBox]::Show("Did we login to Atech correctly?" , "Status" , 4)
			if($b -eq "No")
			{
				exit
			}
			if ((Test-Path $rootpath\AtechIgnitionLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechIgnitionLogin.iim
			}
			
			#############################################
			#Move CSV to correct Location				#
			#############################################
	
			Move-Item -path $DownloadDir\AtechIgnitionASN.csv -destination $SaveDir\CSV\$Date-Ignition.csv
			$Shipment = Import-CSV "$SaveDir\CSV\$Date-Ignition.csv"
			if ((Test-Path $DownloadDir\AtechIgnitionASN.csv) -eq $true)
			{
				[System.Windows.Forms.MessageBox]::Show("After this point you will need to download a new CSV if there are any errors") | Out-Null
				Remove-Item $DownloadDir\AtechIgnitionASN.csv
			}
			
			#################################
			#Begin Building IIM				#
			#################################
			
			Add-Content $rootpath\AtechIgnitionASN.iim "VERSION BUILD=7401110 RECORDER=FX"
			Add-Content $rootpath\AtechIgnitionASN.iim "`nTAB T=1"
			Add-Content $rootpath\AtechIgnitionASN.iim "URL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"

			#####################################################
			#Determine if this is for the first line or not.	#
			#####################################################
			
			if ($i -lt 1)
			{
				$i = 0
			}
			
			#####################################################
			#Insert PO and Head to Creation Form				#
			#Loop to move from one PO to the next.				#
			#####################################################
			
			while ($i -le $Shipment.length-1)
			{
				$PONumber = $Shipment[$i].PONUMBER
				Add-Content $rootpath\AtechIgnitionASN.iim "`n'Start at ASN Creation Screen. User must be logged in!
				`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
				Add-Content $rootpath\AtechIgnitionASN.iim "`n'Fill in PO#"
				Add-Content $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
				Add-Content $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
				Add-Content $rootpath\AtechIgnitionASN.iim "`n'Fill in the Quantity Shipped"
				$q = 1
				
				#################################################
				#Fill in Item Quantity							#
				#################################################
				
				while ($Shipment[$i].PONUMBER -eq $PONumber)
				{
					$QTYShipped = $Shipment[$i].QTYSHIPPED
					Add-Content $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=ID:ASNQty$q CONTENT=$QTYShipped"
					
					#################################################
					#Line Complete?									#
					#################################################
					
					$LineComplete = $Shipment[$i].LINECOMPLETE
				
					if($LineComplete -eq "Y")
					{
						Add-Content $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:frmASNCreate ATTR=ID:checkbox_$q CONTENT=YES"
					}
		
					$i++
					$q++
				}
				
				#################################################
				#Determine Carrier Code							#
				#################################################
				
				if ($Shipment[$i-1].SHIPPERNAME -eq "UPS  - Ground")
				{
					$Carrier = "%0001"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Residential ")
				{    
					$Carrier = "%0001"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "FedEx  - Ground")
				{
					$Carrier = "%0005"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label Res'd")
				{
					$Carrier = "%0040"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Blue Label Res'd")
				{
					$Carrier = "%0041"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Blue Label ")
				{
					$Carrier = "%0041"
				}elseif($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label ")
				{
					$Carrier = "%0040"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label Saturday Delivery")
				{
					$Carrier = "%0043"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq '"Roadrunner"')
				{
					$Carrier = "%0034"
				}elseif($Shipment[$i-1].SHIPPERNAME -eq '"UPS Freight LTL Standard"')
				{
					$Carrier = "%0055"
				}else
				{
					$Carrier = "%0023"
					$Comment = $Shipment[$i-1].SHIPPERNAME
					Add-Content $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=TEXTAREA FORM=NAME:frmASNCreate ATTR=NAME:Comments CONTENT=$Comment"
				}
				Add-Content $rootpath\AtechIgnitionASN.iim "`n'Choose Carrier `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:CarrierID CONTENT=$Carrier"
		
				$DateShipped = $Shipment[$i-1].SHIPDATE
				Add-Content $rootpath\AtechIgnitionASN.iim "`n'Shipment Date `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:Arrival CONTENT=%$DateShipped"
			
				$TrackingNumber = $Shipment[$i-1].TRACKINGNO
				Add-Content $rootpath\AtechIgnitionASN.iim "`n'Tracking Number Insertion`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:BOL CONTENT=$TrackingNumber"
		
				#########################################
				#Insert InvoiceNumber					#
				#########################################
				
				$InvoiceNumber = $Shipment[$i-1].INVOICENUMBER
				Add-Content $rootpath\AtechIgnitionASN.iim "`n'Add InvoiceNumber to Invoice Field`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:VendorReference CONTENT=$InvoiceNumber"
				Add-Content $rootpath\AtechIgnitionASN.iim "WAIT SECONDS = 5"
				Add-Content $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:frmASNCreate ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
				Add-Content $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmApprove ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
		
				#########################################
				#Save ASN as a Webpage					#
				#########################################
				Add-Content $rootpath\AtechIgnitionASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\ FILE=$PONumber-{{!NOW:mm-dd-yyyy-hhnnss}}"
	
			}
			
			#########################################
			#End Loop Here							#
			#########################################
            
			#########################################
			#Logout									#
			#########################################
			
			Add-Content $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=A ATTR=TXT:Log<SP>Out"
			
			#####################################################
			#Run the Macro in Firefox, Firefox must be started!	#
			#####################################################
    
			if ((Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe") -eq $true)
			{
				$cmdLine = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
			}else
			{
				$cmdLine = "C:\Program Files\Mozilla Firefox\firefox.exe"
			}
			$args = "imacros://run/?m=AtechIgnitionASN.iim"
			start-process $cmdLine $args
			Get-Process | ? {$_.Name -like "firefox"} | %{$_.Close()}
        }
        
        $i = $null
		$b = $null
		$a = $null
		
    }

	$Shipment = $null
	
		#########################################
		#Start Exhaust block of code			#
		#########################################
		
	if ($Division -eq "Exhaust")
	{
		if ((Test-Path $DownloadDir\AtechExhaustASN.csv) -eq $true)
		{
		
			#################################
			#Remove old files				#
			#################################
			
			if ((Test-Path $rootpath\AtechExhaustASN.iim) -eq $true)
			{    
				Remove-Item $rootpath\AtechExhaustASN.iim
			}
			if ((Test-Path $rootpath\AtechExhaustLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechExhaustLogin.iim
			}
			$username = Read-Host 'Please input your Username for Exhaust'
			$pass = Read-Host 'Please input your password?' -AsSecureString
			
			$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))
			
			#############################################
			#Build Login Macro							#
			#############################################
			
			Add-Content $rootpath\AtechExhaustLogin.iim "TAB T=1 `nURL GOTO=https://b2b.atechmotorsports.com/default.asp"
			Add-Content $rootpath\AtechExhaustLogin.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmLogin ATTR=NAME:UserName CONTENT=$username"
			Add-Content $rootpath\AtechExhaustLogin.iim "SET !ENCRYPTION NO"
			Add-Content $rootpath\AtechExhaustLogin.iim "TAG POS=1 TYPE=INPUT:PASSWORD FORM=NAME:frmLogin ATTR=NAME:Password CONTENT=$password"
			Add-Content $rootpath\AtechExhaustLogin.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmLogin ATTR=VALUE:Go"
			$password = $null
			
			
			#############################################
			#Run Login Macro							#
			#############################################
			
			if ((Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe") -eq $true)
			{
				$cmdLine = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
			}else
			{
				$cmdLine = "C:\Program Files\Mozilla Firefox\firefox.exe"
			}
			$args = "imacros://run/?m=AtechExhaustLogin.iim"
			start-process $cmdLine $args
			Get-Process | ? {$_.Name -like "firefox"} | %{$_.Close()}
    
			#############################################
			#Ready to move to the Automation?			#
			#############################################
			
			#$a = new-object -comobject wscript.shell
			#$b = $a.popup("Did the script log on to Atech?",0,"Test Message Box",1)
			#if($b -eq 2)
			#{
			#	exit
			#}

			$b = [System.Windows.Forms.MessageBox]::Show("Did we login to Atech correctly?" , "Status" , 4)
			if($b -eq "No")
			{
				exit
			}
			if ((Test-Path $rootpath\AtechExhaustLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechExhaustLogin.iim
			}
			
			#############################################
			#Move CSV to correct Location				#
			#############################################
			
			Move-Item -path $DownloadDir\AtechExhaustASN.csv -destination $SaveDir\CSV\$Date-Exhaust.csv
			$Shipment = Import-CSV "$SaveDir\CSV\$Date-Exhaust.csv"
			if ((Test-Path $DownloadDir\AtechExhaustASN.csv) -eq $true)
			{
				[System.Windows.Forms.MessageBox]::Show("After this point you will need to download a new CSV if there are any errors") | Out-Null
				Remove-Item $DownloadDir\AtechExhaustASN.csv
			}
		
			#############################################
			#Begin Building IIM							#
			#############################################
			
			Add-Content $rootpath\AtechExhaustASN.iim "VERSION BUILD=7401110 RECORDER=FX"
			Add-Content $rootpath\AtechExhaustASN.iim "`nTAB T=1"
			Add-Content $rootpath\AtechExhaustASN.iim "URL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
			
			#####################################################
			#Determine if this is for the first line or not.	#
			#####################################################
			
			if ($i -lt 1)
			{
				$i = 0
			}

			#####################################################
			#Insert PO and Head to Creation Form				#
			#Loop to move from one PO to the next.				#
			#####################################################
			
			while ($i -le $Shipment.length-1)
			{
				$PONumber = $Shipment[$i].PONUMBER
				Add-Content $rootpath\AtechExhaustASN.iim "`n'Start at ASN Creation Screen. User must be logged in!
				`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
				Add-Content $rootpath\AtechExhaustASN.iim "`n'Fill in PO#"
				Add-Content $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
				Add-Content $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
				Add-Content $rootpath\AtechExhaustASN.iim "`n'Fill in the Quantity Shipped"
				$q = 1
			
				#################################################
				#Fill in Item Quantity							#
				#################################################
				
				while ($Shipment[$i].PONUMBER -eq $PONumber)
				{
					$QTYShipped = $Shipment[$i].QTYSHIPPED
    
					Add-Content $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=ID:ASNQty$q CONTENT=$QTYShipped"
					
					#############################################
					#Line Complete?								#
					#############################################
					
					$LineComplete = $Shipment[$i].LINECOMPLETE
				
					if($LineComplete -eq "Y")
					{
						Add-Content $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:frmASNCreate ATTR=ID:checkbox_$q CONTENT=YES"
					}
		
					$i++
					$q++
				}
				
				#################################################
				#Determine Carrier Code							#
				#################################################
				
				if ($Shipment[$i-1].SHIPPERNAME -eq "UPS  - Ground")
				{
					$Carrier = "%0001"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Residential ")
				{    
					$Carrier = "%0001"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "FedEx  - Ground")
				{
					$Carrier = "%0005"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label Res'd")
				{
					$Carrier = "%0040"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Blue Label Res'd")
				{
					$Carrier = "%0041"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Blue Label ")
				{
					$Carrier = "%0041"
				}elseif($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label ")
				{
					$Carrier = "%0040"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label Saturday Delivery")
				{
					$Carrier = "%0043"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq '"Roadrunner"')
				{
					$Carrier = "%0034"
				}elseif($Shipment[$i-1].SHIPPERNAME -eq '"UPS Freight LTL Standard"')
				{
					$Carrier = "%0055"
				}else
				{
					$Carrier = "%0023"
					$Comment = $Shipment[$i-1].SHIPPERNAME
					Add-Content $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=TEXTAREA FORM=NAME:frmASNCreate ATTR=NAME:Comments CONTENT=$Comment"
				}
				Add-Content $rootpath\AtechExhaustASN.iim "`n'Choose Carrier `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:CarrierID CONTENT=$Carrier"
		
				$DateShipped = $Shipment[$i-1].SHIPDATE
				Add-Content $rootpath\AtechExhaustASN.iim "`n'Shipment Date `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:Arrival CONTENT=%$DateShipped"
			
				$TrackingNumber = $Shipment[$i-1].TRACKINGNO
				Add-Content $rootpath\AtechExhaustASN.iim "`n'Tracking Number Insertion`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:BOL CONTENT=$TrackingNumber"
				
				#########################################
				#Insert InvoiceNumber					#
				#########################################
				
				$InvoiceNumber = $Shipment[$i-1].INVOICENUMBER
				Add-Content $rootpath\AtechExhaustASN.iim "`n'Add InvoiceNumber to Invoice Field`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:VendorReference CONTENT=$InvoiceNumber"
				Add-Content $rootpath\AtechExhaustASN.iim "WAIT SECONDS = 5"
				Add-Content $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:frmASNCreate ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
				Add-Content $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmApprove ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
		
				#########################################
				#Save ASN as a Webpage					#
				#########################################
				
				Add-Content $rootpath\AtechExhaustASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\ FILE=$PONumber-{{!NOW:mm-dd-yyyy-hhnnss}}"
				

			}
			
			#########################################
			#End Loop Here							#
			#########################################

			#########################################
			#Logout									#
			#########################################
			
			Add-Content $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=A ATTR=TXT:Log<SP>Out"
			
			#################################################################
			#Run the Macro in Firefox, Firefox must be started				#
			#################################################################
    
			if ((Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe") -eq $true)
			{
				$cmdLine = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
			}else
			{
				$cmdLine = "C:\Program Files\Mozilla Firefox\firefox.exe"
			}
			$args = "imacros://run/?m=AtechExhaustASN.iim"
			start-process $cmdLine $args
			Get-Process | ? {$_.Name -like "firefox"} | %{$_.Close()}
        }



        $i = $null
		$a = $null
		$b = $null
    }

	$Shipment = $null

	
	
		#############################################
		#Start ExhPrivateLabel block of code		#
		#############################################
		
	if ($Division -eq "PrivateLabel")
	{
		if ((Test-Path $DownloadDir\AtechExhPrivateLabelASN.csv) -eq $true)
		{
		
			#############################
			#Remove old files			#
			#############################
			
			if ((Test-Path $rootpath\AtechExhPrivateLabelASN.iim) -eq $true)
			{    
				Remove-Item $rootpath\AtechExhPrivateLabelASN.iim
			}
			if ((Test-Path $rootpath\AtechExhPrivateLabelLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechExhPrivateLabelLogin.iim
			}
			$username = Read-Host 'Please input your Username for Private Label'
			$pass = Read-Host 'Please input your password?' -AsSecureString
			
			$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))
			
			#############################################
			#Build Login Macro							#
			#############################################
			
			Add-Content $rootpath\AtechExhPrivateLabelLogin.iim "TAB T=1 `nURL GOTO=https://b2b.atechmotorsports.com/default.asp"
			Add-Content $rootpath\AtechExhPrivateLabelLogin.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmLogin ATTR=NAME:UserName CONTENT=$username"
			Add-Content $rootpath\AtechExhPrivateLabelLogin.iim "SET !ENCRYPTION NO"
			Add-Content $rootpath\AtechExhPrivateLabelLogin.iim "TAG POS=1 TYPE=INPUT:PASSWORD FORM=NAME:frmLogin ATTR=NAME:Password CONTENT=$password"
			Add-Content $rootpath\AtechExhPrivateLabelLogin.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmLogin ATTR=VALUE:Go"
			$password = $null
			
			
			#############################################
			#Run Login Macro							#
			#############################################
			
			if ((Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe") -eq $true)
			{
				$cmdLine = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
			}else
			{
				$cmdLine = "C:\Program Files\Mozilla Firefox\firefox.exe"
			}
			$args = "imacros://run/?m=AtechExhPrivateLabelLogin.iim"
			start-process $cmdLine $args
			Get-Process | ? {$_.Name -like "firefox"} | %{$_.Close()}
    
			#############################################
			#Ready to move to the Automation?			#
			#############################################
			
			#$a = new-object -comobject wscript.shell
			#$b = $a.popup("Did the script log on to Atech?",0,"Test Message Box",4)
			#if($b -eq 2)
			#{
			#	exit
			#}
			$b = [System.Windows.Forms.MessageBox]::Show("Did we login to Atech correctly?" , "Status" , 4)
			if($b -eq "No")
			{
				exit
			}
			if ((Test-Path $rootpath\AtechExhPrivateLabelLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechExhPrivateLabelLogin.iim
			}
			
			#############################################
			#Move CSV to correct Location				#
			#############################################
			
			Move-Item -path $DownloadDir\AtechExhPrivateLabelASN.csv -destination $SaveDir\CSV\$Date-ExhPrivateLabel.csv
			$Shipment = Import-CSV "$SaveDir\CSV\$Date-ExhPrivateLabel.csv"
			if ((Test-Path $DownloadDir\AtechExhPrivateLabelASN.csv) -eq $true)
			{
				[System.Windows.Forms.MessageBox]::Show("After this point you will need to download a new CSV if there are any errors") | Out-Null
				Remove-Item $DownloadDir\AtechExhPrivateLabelASN.csv
			}
		
			#################################
			#Begin Building IIM				#
			#################################
			
			Add-Content $rootpath\AtechExhPrivateLabelASN.iim "VERSION BUILD=7401110 RECORDER=FX"
			Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`nTAB T=1"
			Add-Content $rootpath\AtechExhPrivateLabelASN.iim "URL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"


			#####################################################
			#Determine if this is for the first line or not.	#
			#####################################################
			
			if ($i -lt 1)
			{
				$i = 0
			}
			
			#################################################
			#Insert PO and Head to Creation Form			#
			#Loop to move from one PO to the next.			#
			#################################################
			
			while ($i -le $Shipment.length-1)
			{
				$PONumber = $Shipment[$i].PONUMBER
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`n'Start at ASN Creation Screen. User must be logged in!
				`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`n'Fill in PO#"
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`n'Fill in the Quantity Shipped"
				$q = 1
				
				#################################
				#Fill in Item Quantity			#
				#################################
				
				while ($Shipment[$i].PONUMBER -eq $PONumber)
				{
					$QTYShipped = $Shipment[$i].QTYSHIPPED
		
					Add-Content $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=ID:ASNQty$q CONTENT=$QTYShipped"
					
					#################################
					#Line Complete?					#
					#################################
					
					$LineComplete = $Shipment[$i].LINECOMPLETE
				
					if($LineComplete -eq "Y")
					{
						Add-Content $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:frmASNCreate ATTR=ID:checkbox_$q CONTENT=YES"
					}
			
					$i++
					$q++
				}
				
				#####################################
				#Determine Carrier Code				#
				#####################################
				
				if ($Shipment[$i-1].SHIPPERNAME -eq "UPS  - Ground")
				{
					$Carrier = "%0001"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Residential ")
				{    
					$Carrier = "%0001"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "FedEx  - Ground")
				{
					$Carrier = "%0005"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label Res'd")
				{
					$Carrier = "%0040"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Blue Label Res'd")
				{
					$Carrier = "%0041"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Blue Label ")
				{
					$Carrier = "%0041"
				}elseif($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label ")
				{
					$Carrier = "%0040"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label Saturday Delivery")
				{
					$Carrier = "%0043"
				}elseif ($Shipment[$i-1].SHIPPERNAME -eq '"Roadrunner"')
				{
					$Carrier = "%0034"
				}elseif($Shipment[$i-1].SHIPPERNAME -eq '"UPS Freight LTL Standard"')
				{
					$Carrier = "%0055"
				}else
				{
					$Carrier = "%0023"
					$Comment = $Shipment[$i-1].SHIPPERNAME
		
					Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=TEXTAREA FORM=NAME:frmASNCreate ATTR=NAME:Comments CONTENT=$Comment"
				}
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`n'Choose Carrier `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:CarrierID CONTENT=$Carrier"
		
				$DateShipped = $Shipment[$i-1].SHIPDATE
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`n'Shipment Date `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:Arrival CONTENT=%$DateShipped"
			
				$TrackingNumber = $Shipment[$i-1].TRACKINGNO
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`n'Tracking Number Insertion`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:BOL CONTENT=$TrackingNumber"
		
				#########################################
				#Insert InvoiceNumber					#
				#########################################
				
				$InvoiceNumber = $Shipment[$i-1].INVOICENUMBER
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`n'Add InvoiceNumber to Invoice Field`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:VendorReference CONTENT=$InvoiceNumber"
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "WAIT SECONDS = 5"
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:frmASNCreate ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmApprove ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
		
				#########################################
				#Save ASN as a Webpage					#
				#########################################
				
				Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\ FILE=$PONumber-{{!NOW:mm-dd-yyyy-hhnnss}}"
				

			}
			
			#########################################
			#End Loop Here							#
			#########################################

			#########################################
			#Logout									#
			#########################################
			Add-Content $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=A ATTR=TXT:Log<SP>Out"
			
			#################################################################
			#Run the Macro in Firefox, Firefox must be started.				#
			#################################################################
    
			if ((Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe") -eq $true)
			{
				$cmdLine = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
			}else
			{
				$cmdLine = "C:\Program Files\Mozilla Firefox\firefox.exe"
			}
			$args = "imacros://run/?m=AtechExhPrivateLabelASN.iim"
			start-process $cmdLine $args
			Get-Process | ? {$_.Name -like "firefox"} | %{$_.Close()}
        }


		#########################################
        #For Testing Purposes Only				#
        #########################################

        $i = $null
		$a = $null
		$b = $null
    }

	$Shipment = $null
	
	$Continue = [System.Windows.Forms.MessageBox]::Show("Would you like to continue?" , "Status" , 4)
}