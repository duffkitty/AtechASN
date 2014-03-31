#########################################################
#ASN Generation Script for Atech						
#Written by Anthony Sinatra								
#Written for PerTronix								
#Build 0.3.6 - Beta										
#########################################################

#########################################################
#Notes
#########################################################

#########################################################
#Set Required Paths										
#														
#														
#Setting the Continue Variable. If this isn't set code	
#will exit!												
#########################################################

$Continue = "Yes"
$rootpath = "$home\Documents\iMacros\Macros"
$SaveDir = "\\snap\share1\ASN\Atech"
$CSVDir = "\\snap\share1\ASN\Atech\CSV"
$DownloadDir = "$home\Downloads"
$Date = get-date -uformat %m%d%Y

#########################################################
#Make Directory if it doesn't exist
#########################################################

if((Test-Path $SaveDir\$Date) -eq $false)
{
	md $SaveDir\$Date
}

#########################################################
#Functions!
#########################################################

function AreArraysEqual($a1, $a2) {
    if ($a1 -isnot [array] -or $a2 -isnot [array])
	{ 
      throw "Both inputs must be an array"
    }
    if ($a1.Rank -ne $a2.Rank)
	{ 
      return $false 
    }
    if ([System.Object]::ReferenceEquals($a1, $a2))
	{
      return $true
    }
    for ($r = 0; $r -lt $a1.Rank; $r++)
	{
      if ($a1.GetLength($r) -ne $a2.GetLength($r))
	  {
            return $false
      }
    }
    $enum1 = $a1.GetEnumerator()
    $enum2 = $a2.GetEnumerator()   

    while ($enum1.MoveNext() -and $enum2.MoveNext())
	{
      if ($enum1.Current -ne $enum2.Current)
	  {
            return $false
      }
    }
    return $true
}

#########################################################
#Remove Old files from older versions					
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
#Begin Code!											
#########################################################

while($Continue -eq "Yes")
{
	#####################################################
	#region functions									
	#####################################################
	
	#####################################################
	# load WinForms										
	#####################################################
	
	[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	
	#####################################################
	# create form										
	#####################################################
	
	$form = New-Object Windows.Forms.Form
	$form.text = "Choose Division"
	$form.top = 10
	$form.left = 10
	$form.height = 250
	$form.width = 350

	#####################################################
	# create label										
	#####################################################
	
	$label = New-Object Windows.Forms.Label
	$label.text = "Which Division would you like to create ASNs for?"
	$label.height = 275
	$label.width = 100
	$label.top = 2
	$label.left = 25
	$form.controls.add($label)
	
	#####################################################
	# create button										
	#####################################################
	
	$button = New-Object Windows.Forms.Button
	$button.text = "Run!"
	$button.height = 40
	$button.width = 70
	$button.top = 150
	$button.left = 150
	$form.controls.add($button)
	
	#####################################################
	# create radiobutton								
	#####################################################
	
	$RadioButton = New-Object Windows.Forms.radiobutton
	$RadioButton.text = "Ignition"
	$RadioButton.height = 20
	$RadioButton.width = 150
	$RadioButton.top = 2
	$RadioButton.left = 150
	$form.controls.add($RadioButton)

	#####################################################
	# create radiobutton1								
	#####################################################
	
	$radiobutton1 = New-Object Windows.Forms.radiobutton
	$RadioButton1.text = "Exhaust"
	$RadioButton1.height = 20
	$RadioButton1.width = 150
	$RadioButton1.top = 30
	$RadioButton1.left =150
	$form.controls.add($RadioButton1)

	#####################################################
	# create radiobutton2								
	#####################################################
	
	$radiobutton2 = New-Object Windows.Forms.radiobutton
	$RadioButton2.text = "Private Label"
	$RadioButton2.height = 20
	$RadioButton2.width = 150
	$RadioButton2.top = 58
	$RadioButton2.left =150
	$form.controls.add($RadioButton2)

	#####################################################
	# create event handler for button					
	#####################################################
	
	$event = {
		if($radiobutton.checked){$Global:Division = "Ignition"}
		if($radiobutton1.checked){$Global:Division = "Exhaust"}
		if($radiobutton2.checked){$Global:Division = "PrivateLabel"}
		$form.Close()
	}

	#####################################################
	# attach event handler								
	#####################################################
	
	$button.Add_Click($event)

	#####################################################
	# attach controls to form							
	#####################################################
	
	$form.controls.add($button)
	$form.controls.add($label)
	$form.controls.add($textbox)

	$form.showdialog()



	#####################################################
	#Start Ignition block of code						
	#####################################################
	
	if ($Division -eq "Ignition")
	{
		#########################################################
		#Create Error Array
		#########################################################
		
		$Errors = @()

		if ((Test-Path $DownloadDir\AtechIgnitionASN.csv) -eq $true)
		{
		
			#############################################
			#Remove old files							
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
			#Login information							
			#############################################

			$username = Read-Host 'Please input your Username for Ignition'
			$pass = Read-Host 'Please input your password?' -AsSecureString
			
			#############################################
			#Convert password to plain text variable	
			#############################################
			
			$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))
			
			#############################################
			#Build Login Macro 							
			#############################################
			
			Add-Content -Encoding UTF8 $rootpath\AtechIgnitionLogin.iim "TAB T=1 `nURL GOTO=https://b2b.atechmotorsports.com/default.asp"
			Add-Content -Encoding UTF8 $rootpath\AtechIgnitionLogin.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmLogin ATTR=NAME:UserName CONTENT=$username"
			Add-Content -Encoding UTF8 $rootpath\AtechIgnitionLogin.iim "SET !ENCRYPTION NO"
			Add-Content -Encoding UTF8 $rootpath\AtechIgnitionLogin.iim "TAG POS=1 TYPE=INPUT:PASSWORD FORM=NAME:frmLogin ATTR=NAME:Password CONTENT=$password"
			Add-Content -Encoding UTF8 $rootpath\AtechIgnitionLogin.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmLogin ATTR=VALUE:Go"
			Add-Content -Encoding UTF8 $rootpath\AtechIgnitionLogin.iim "SAVEAS TYPE=TXT FOLDER=$CSVDir FILE=login.txt"
			$password = $null
			
			
			#############################################
			#Run Login Macro							
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
    			
			while (!(Test-Path $CSVDir\login.txt))
			{
				#Wait Until exists!
			}
			
			
			$b = [System.Windows.Forms.MessageBox]::Show("Did we login to Atech correctly?" , "Status" , 4)
			if($b -eq "No")
			{
				exit
			}
			
			if ((Test-Path $CSVDir\login.txt) -eq $true)
			{
				Remove-Item $CSVDir\login.txt
			}
			
			if ((Test-Path $rootpath\AtechIgnitionLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechIgnitionLogin.iim
			}
			
			#############################################
			#Move CSV to correct Location				
			#############################################
	
			Move-Item -path $DownloadDir\AtechIgnitionASN.csv -destination $SaveDir\CSV\$Date-Ignition.csv
			$Shipment = Import-CSV "$SaveDir\CSV\$Date-Ignition.csv"
			if ((Test-Path $DownloadDir\AtechIgnitionASN.csv) -eq $true)
			{
				[System.Windows.Forms.MessageBox]::Show("After this point you will need to download a new CSV if there are any errors") | Out-Null
				Remove-Item $DownloadDir\AtechIgnitionASN.csv
			}
			
			#####################################################
			#Determine if this is for the first line or not.	
			#####################################################
			
			if ($i -lt 1)
			{
				$i = 0
			}
			
			#########################################################
			#Loop Beginning so that script knows where to stop from
			#CSV Length
			#########################################################
			
			
			while ($i -le $Shipment.length-1)
			{			

				#####################################################
				#Begin Checking
				#####################################################
				
				$c = $i
				$v = $i
				$PONumber = $Shipment[$i].PONUMBER
				$POLineComplete = $PONumber
				$LineComplete = @()
				$ProdCheck = @()
				
				###################################################################
				#Build 2 Arrays. One for Checking if the PO is shipping complete
				#The other is to store the product numbers if it is not complete
				###################################################################
				
				while($POLineComplete -eq $PONumber)
				{
					$LineComplete = $LineComplete + $Shipment[$c].LINECOMPLETE
					$ProdCheck = $ProdCheck + $Shipment[$c].PRODNUM
					$c++
					$v++
					$POLineComplete = $Shipment[$v].PONUMBER
				}
				
				if($LineComplete -notcontains "N")
				{
					if((Test-Path $rootpath\AtechIgnitionASN.iim) -eq $true)
					{
						Remove-Item $rootpath\AtechIgnitionASN.iim
					}
					##############################################
					#Ship complete Piece of CAKE!
					##############################################
					
					#####################################################
					#Insert PO and Head to Creation Form				
					#####################################################
				
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "TAB T=1"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Start at ASN Creation Screen.
					`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "SET !ERRORIGNORE YES"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "SET !TIMEOUT_STEP 0"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Fill in PO#"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Fill in the Quantity Shipped"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:frmASNCreate ATTR=NAME:CheckAllBoxes CONTENT=YES"
				
				
					#################################################
					#Determine Carrier Code							
					#################################################
					
					if ($Shipment[$i].SHIPPERNAME -eq "UPS  - Ground")
					{
						$Carrier = "%0001"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Residential ")
					{    
						$Carrier = "%0001"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "FedEx  - Ground")
					{
						$Carrier = "%0005"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Red Label Res'd")
					{
						$Carrier = "%0040"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Blue Label Res'd")
					{
						$Carrier = "%0041"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Blue Label ")
					{
						$Carrier = "%0041"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Red Label ")
					{
						$Carrier = "%0040"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Red Label Saturday Delivery")
					{
						$Carrier = "%0043"
					}elseif ($Shipment[$i].SHIPPERNAME -eq '"Roadrunner"')
					{
						$Carrier = "%0034"
					}elseif ($Shipment[$i-1].SHIPPERNAME -eq "Roadrunner")
					{
						$Carrier = "%0034"
					}elseif ($Shipment[$i].SHIPPERNAME -eq '"UPS Freight LTL Standard"')
					{
						$Carrier = "%0055"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Freight LTL Standard")
					{
						$Carrier = "%0055"
					}else
					{
						$Carrier = "%0023"
						$Comment = $Shipment[$i].SHIPPERNAME
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=TEXTAREA FORM=NAME:frmASNCreate ATTR=NAME:Comments CONTENT=`"$Comment`""
					}
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Choose Carrier `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:CarrierID CONTENT=$Carrier"
			
					$DateShipped = $Shipment[$i].SHIPDATE
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Shipment Date `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:Arrival CONTENT=%$DateShipped"
				
					$TrackingNumber = $Shipment[$i].TRACKINGNO
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Tracking Number Insertion`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:BOL CONTENT=$TrackingNumber"
			
					#########################################
					#Insert InvoiceNumber					
					#########################################
					
					$InvoiceNumber = $Shipment[$i].INVOICENUMBER
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Add InvoiceNumber to Invoice Field`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:VendorReference CONTENT=$InvoiceNumber"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:frmASNCreate ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmApprove ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
			
					#########################################
					#Save ASN as a Webpage					
					#########################################
					if((Test-Path $SaveDir\$Date\$PONumber-$Date.htm) -eq $true)
					{
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date-extra"
					}
					else
					{
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date"
					}
					
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "SAVEAS TYPE=TXT FOLDER=$SaveDir\ FILE=done.txt"
					
					$i = $v
					
				}
				
				else
				{
					##############################################
					#Continue checking!
					##############################################
					$c = $i
					$s = "tblDataRegRowBorder"
                    $p = 3
					$PONumberLines = $Shipment[$i].PONUMBER
					if((Test-Path $rootpath\AtechIgnitionCheck.iim) -eq $true)
					{
						Remove-Item $rootpath\AtechIgnitionCheck.iim
					}
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "TAB T=1"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "`n'Start at ASN Creation Screen.
					`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "SET !EXTRACT_TEST_POPUP NO"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "SET !ERRORIGNORE YES"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "SET !TIMEOUT_STEP 0"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "`n'Fill in PO#"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
			
                    while($PONumberLines -eq $PONumber)
					{
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "TAG POS=$p TYPE=TD FORM=ID:frmASNCreate ATTR=CLASS:$s EXTRACT=TXT"
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "SAVEAS TYPE=EXTRACT FOLDER=$SaveDir FILE=AtechCheck-$PONumber.csv"
						if($s -eq "tblDataShadeRowBorder")
						{
							$p = $p + 8
						}
						if($s -eq "tblDataRegRowBorder")
						{
							$s = "tblDataShadeRowBorder"
						}
						elseif($s -eq "tblDataShadeRowBorder")
						{
							$s = "tblDataRegRowBorder"
						}
						$c++
						$PONumberLines = $Shipment[$c].PONUMBER
					}
					Add-Content -Encoding UTF8 $rootpath\AtechIgnitionCheck.iim "SAVEAS TYPE=TXT FOLDER=$SaveDir FILE=done.txt"
					
					if ((Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe") -eq $true)
					{
						$cmdLine = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
					}else
					{
						$cmdLine = "C:\Program Files\Mozilla Firefox\firefox.exe"
					}
					$args = "imacros://run/?m=AtechIgnitionCheck.iim"
					start-process $cmdLine $args
					Get-Process | ? {$_.Name -like "firefox"} | %{$_.Close()}
					
					
					while (!(Test-Path $SaveDir\done.txt))
					{
						#Wait Until exists!
					}
					Remove-Item $SaveDir\done.txt				
					$CheckPNX = Get-Content $SaveDir\AtechCheck-$PONumber.csv
					$Check = $CheckPNX | ForEach-Object {$_ -replace '"', ""} | ForEach-Object { $_.Remove(0,4) }
					if((AreArraysEqual $Check $ProdCheck) -eq $false)
					{
					
						#####################################################
						#Check failed, bail out!
						#####################################################
					
						$Errors = $Errors + "$PONumber has a mismatched line. Either an incorrect part number was entered or it was entered out of order."
                        $i = $v
					}	
					elseif((AreArraysEqual $Check $ProdCheck) -eq $true)
					{
						If((Test-Path $rootpath\AtechIgnitionASN.iim) -eq $true)
						{
							Remove-Item $rootpath\AtechIgnitionASN.iim
						}
						########################################################
						#Check is ok! Continue to create ASN
						########################################################
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Start at ASN Creation Screen. User must be logged in!
						`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Fill in PO#"
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Fill in the Quantity Shipped"
						$q = 1
					
						#################################################
						#Fill in Item Quantity							
						#################################################
						
						while ($Shipment[$i].PONUMBER -eq $PONumber)
						{	
	
							#################################################
							#Line Complete?									
							#################################################
	
							$LineComplete = $Shipment[$i].LINECOMPLETE
							if($LineComplete -eq "Y")
							{
								Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:frmASNCreate ATTR=ID:checkbox_$q CONTENT=YES"
							}
							else
							{
								$QTYShipped = $Shipment[$i].QTYSHIPPED
								Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=ID:ASNQty$q CONTENT=$QTYShipped"
							}
							
						
				
							$i++
							$q++
						}
						
						#################################################
						#Determine Carrier Code							
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
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label ")
						{
							$Carrier = "%0040"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label Saturday Delivery")
						{
							$Carrier = "%0043"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq '"Roadrunner"')
						{
							$Carrier = "%0034"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq "Roadrunner")
						{
							$Carrier = "%0034"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq '"UPS Freight LTL Standard"')
						{
							$Carrier = "%0055"
						}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Freight LTL Standard")
						{
							$Carrier = "%0055"
						}else
						{
							$Carrier = "%0023"
							$Comment = $Shipment[$i-1].SHIPPERNAME
							Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=TEXTAREA FORM=NAME:frmASNCreate ATTR=NAME:Comments CONTENT=`"$Comment`""
						}
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Choose Carrier `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:CarrierID CONTENT=$Carrier"
				
						$DateShipped = $Shipment[$i-1].SHIPDATE
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Shipment Date `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:Arrival CONTENT=%$DateShipped"
					
						$TrackingNumber = $Shipment[$i-1].TRACKINGNO
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Tracking Number Insertion`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:BOL CONTENT=$TrackingNumber"
				
						#########################################
						#Insert InvoiceNumber					
						#########################################
						
						$InvoiceNumber = $Shipment[$i-1].INVOICENUMBER
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`n'Add InvoiceNumber to Invoice Field`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:VendorReference CONTENT=$InvoiceNumber"
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "WAIT SECONDS = 5"
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:frmASNCreate ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmApprove ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
				
						#########################################
						#Save ASN as a Webpage					
						#########################################
						if((Test-Path $SaveDir\$Date\$PONumber-$Date) -eq $true)
						{
							Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date-extra"
						}
						else
						{
							Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date"
						}
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "SAVEAS TYPE=TXT FOLDER=$SaveDir\ FILE=done.txt"

					}
				}
				#########################################
				#Logout									
				#########################################
			
				if($i -eq $Shipment.length)
				{
						Add-Content -Encoding UTF8 $rootpath\AtechIgnitionASN.iim "`nTAG POS=1 TYPE=A ATTR=TXT:Log<SP>Out"
				}				
				
				#####################################################
				#Run the Macro in Firefox, Firefox must be started!	
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
				
				while (!(Test-Path $SaveDir\done.txt))
				{
					#Wait Until exists!
				}
				
				$f = Get-Content $SaveDir\done.txt | Select-String "No Results Found" -quiet
				if($f -eq $true)
				{
					$Errors = $Errors + "$PONumber is not a valid PO and had a fatal error"
					Remove-Item "$SaveDir\$Date\$PONumber-$Date"
				}
				$f = $null
				$f = Get-Content $SaveDir\done.txt | Select-String "No items found with search criteria." -quiet
				if($f -eq $true)
				{
					$Errors = $Errors + "$PONumber was probably already finished. If not please run manually."
				}
					
				Remove-Item $SaveDir\done.txt
	
			
				
				$LineComplete = $null
				$ProdCheck = $null
			
			}
				
				

			#########################################
			#End Loop Here							
			#########################################
			if($Errors.length -gt 0)
			{
				[reflection.assembly]::loadwithpartialname('system.windows.forms');
				[system.Windows.Forms.MessageBox]::show($Errors)
				[reflection.assembly]::loadwithpartialname('system.windows.forms');
				[system.Windows.Forms.MessageBox]::show("If you did not catch that, the errors are logged in the ASN folder on the Snap Drive under todays date.")
				
				if($Errors.length -gt 0)
				{
					$Errors | Export-CSV $SaveDir\$Date\Ignition-Errors.csv
				}
				$Errors = $Null
			}
			$Continue = [System.Windows.Forms.MessageBox]::Show("ASN creation for Ignition should now be complete. Would you like to run another division?" , "Status" , 4)
		}
		
		Else
		{
			$b = [System.Windows.Forms.MessageBox]::Show("Please make sure you downloaded the Ignition CSV from Pentaho. Would you like to start over?" , "Status" , 4)
			if($b -eq "No")
			{
				exit
			}
		}
        
        $i = $null
		$b = $null
		$a = $null
		$v = $null
		$Division = $null
    }

	$Shipment = $null

	#####################################################
	#Start Exhaust block of code						
	#####################################################
	
	#####################################################
	#Start Exhaust block of code						
	#####################################################
	
	if ($Division -eq "Exhaust")
	{
		#########################################################
		#Create Error Array
		#########################################################
		
		$Errors = @()

		if ((Test-Path $DownloadDir\AtechExhaustASN.csv) -eq $true)
		{
		
			#############################################
			#Remove old files							
			#############################################
			
			if ((Test-Path $rootpath\AtechExhaustASN.iim) -eq $true)
			{    
				Remove-Item $rootpath\AtechExhaustASN.iim
			}
			if ((Test-Path $rootpath\AtechExhaustLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechExhaustLogin.iim
			}
			
			#############################################
			#Login information							
			#############################################

			$username = Read-Host 'Please input your Username for Exhaust'
			$pass = Read-Host 'Please input your password?' -AsSecureString
			
			#############################################
			#Convert password to plain text variable	
			#############################################
			
			$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))
			
			#############################################
			#Build Login Macro 							
			#############################################
			
			Add-Content -Encoding UTF8 $rootpath\AtechExhaustLogin.iim "TAB T=1 `nURL GOTO=https://b2b.atechmotorsports.com/default.asp"
			Add-Content -Encoding UTF8 $rootpath\AtechExhaustLogin.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmLogin ATTR=NAME:UserName CONTENT=$username"
			Add-Content -Encoding UTF8 $rootpath\AtechExhaustLogin.iim "SET !ENCRYPTION NO"
			Add-Content -Encoding UTF8 $rootpath\AtechExhaustLogin.iim "TAG POS=1 TYPE=INPUT:PASSWORD FORM=NAME:frmLogin ATTR=NAME:Password CONTENT=$password"
			Add-Content -Encoding UTF8 $rootpath\AtechExhaustLogin.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmLogin ATTR=VALUE:Go"
			Add-Content -Encoding UTF8 $rootpath\AtechExhaustLogin.iim "SAVEAS TYPE=TXT FOLDER=$CSVDir FILE=login.txt"
			$password = $null
			
			
			#############################################
			#Run Login Macro							
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
    			
			while (!(Test-Path $CSVDir\login.txt))
			{
				#Wait Until exists!
			}
			
			
			$b = [System.Windows.Forms.MessageBox]::Show("Did we login to Atech correctly?" , "Status" , 4)
			if($b -eq "No")
			{
				exit
			}
			
			if ((Test-Path $CSVDir\login.txt) -eq $true)
			{
				Remove-Item $CSVDir\login.txt
			}
			
			if ((Test-Path $rootpath\AtechExhaustLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechExhaustLogin.iim
			}
			
			#############################################
			#Move CSV to correct Location				
			#############################################
	
			Move-Item -path $DownloadDir\AtechExhaustASN.csv -destination $SaveDir\CSV\$Date-Exhaust.csv
			$Shipment = Import-CSV "$SaveDir\CSV\$Date-Exhaust.csv"
			if ((Test-Path $DownloadDir\AtechExhaustASN.csv) -eq $true)
			{
				[System.Windows.Forms.MessageBox]::Show("After this point you will need to download a new CSV if there are any errors") | Out-Null
				Remove-Item $DownloadDir\AtechExhaustASN.csv
			}
			
			#####################################################
			#Determine if this is for the first line or not.	
			#####################################################
			
			if ($i -lt 1)
			{
				$i = 0
			}
			
			#########################################################
			#Loop Beginning so that script knows where to stop from
			#CSV Length
			#########################################################
			
			
			while ($i -le $Shipment.length-1)
			{			

				#####################################################
				#Begin Checking
				#####################################################
				
				$c = $i
				$v = $i
				$PONumber = $Shipment[$i].PONUMBER
				$POLineComplete = $PONumber
				$LineComplete = @()
				$ProdCheck = @()
				
				###################################################################
				#Build 2 Arrays. One for Checking if the PO is shipping complete
				#The other is to store the product numbers if it is not complete
				###################################################################
				
				while($POLineComplete -eq $PONumber)
				{
					$LineComplete = $LineComplete + $Shipment[$c].LINECOMPLETE
					$ProdCheck = $ProdCheck + $Shipment[$c].PRODNUM
					$c++
					$v++
					$POLineComplete = $Shipment[$v].PONUMBER
				}
				
				if($LineComplete -notcontains "N")
				{
					if((Test-Path $rootpath\AtechExhaustASN.iim) -eq $true)
					{
						Remove-Item $rootpath\AtechExhaustASN.iim
					}
					##############################################
					#Ship complete Piece of CAKE!
					##############################################
					
					#####################################################
					#Insert PO and Head to Creation Form				
					#####################################################
				
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "TAB T=1"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Start at ASN Creation Screen.
					`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "SET !ERRORIGNORE YES"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "SET !TIMEOUT_STEP 0"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Fill in PO#"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Fill in the Quantity Shipped"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:frmASNCreate ATTR=NAME:CheckAllBoxes CONTENT=YES"
				
				
					#################################################
					#Determine Carrier Code							
					#################################################
					
					if ($Shipment[$i].SHIPPERNAME -eq "UPS  - Ground")
					{
						$Carrier = "%0001"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Residential ")
					{    
						$Carrier = "%0001"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "FedEx  - Ground")
					{
						$Carrier = "%0005"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Red Label Res'd")
					{
						$Carrier = "%0040"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Blue Label Res'd")
					{
						$Carrier = "%0041"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Blue Label ")
					{
						$Carrier = "%0041"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Red Label ")
					{
						$Carrier = "%0040"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Red Label Saturday Delivery")
					{
						$Carrier = "%0043"
					}elseif ($Shipment[$i].SHIPPERNAME -eq '"Roadrunner"')
					{
						$Carrier = "%0034"
					}elseif ($Shipment[$i-1].SHIPPERNAME -eq "Roadrunner")
					{
						$Carrier = "%0034"
					}elseif ($Shipment[$i].SHIPPERNAME -eq '"UPS Freight LTL Standard"')
					{
						$Carrier = "%0055"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Freight LTL Standard")
					{
						$Carrier = "%0055"
					}else
					{
						$Carrier = "%0023"
						$Comment = $Shipment[$i].SHIPPERNAME
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=TEXTAREA FORM=NAME:frmASNCreate ATTR=NAME:Comments CONTENT=`"$Comment`""
					}
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Choose Carrier `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:CarrierID CONTENT=$Carrier"
			
					$DateShipped = $Shipment[$i].SHIPDATE
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Shipment Date `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:Arrival CONTENT=%$DateShipped"
				
					$TrackingNumber = $Shipment[$i].TRACKINGNO
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Tracking Number Insertion`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:BOL CONTENT=$TrackingNumber"
			
					#########################################
					#Insert InvoiceNumber					
					#########################################
					
					$InvoiceNumber = $Shipment[$i].INVOICENUMBER
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Add InvoiceNumber to Invoice Field`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:VendorReference CONTENT=$InvoiceNumber"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:frmASNCreate ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmApprove ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
			
					#########################################
					#Save ASN as a Webpage					
					#########################################
					if((Test-Path $SaveDir\$Date\$PONumber-$Date.htm) -eq $true)
					{
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date-extra"
					}
					else
					{
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date"
					}
					
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "SAVEAS TYPE=TXT FOLDER=$SaveDir\ FILE=done.txt"
					
					$i = $v
					
				}
				
				else
				{
					##############################################
					#Continue checking!
					##############################################
					$c = $i
					$s = "tblDataRegRowBorder"
                    $p = 3
					$PONumberLines = $Shipment[$i].PONUMBER
					if((Test-Path $rootpath\AtechExhaustCheck.iim) -eq $true)
					{
						Remove-Item $rootpath\AtechExhaustCheck.iim
					}
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "TAB T=1"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "`n'Start at ASN Creation Screen.
					`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "SET !EXTRACT_TEST_POPUP NO"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "SET !ERRORIGNORE YES"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "SET !TIMEOUT_STEP 0"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "`n'Fill in PO#"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
			
                    while($PONumberLines -eq $PONumber)
					{
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "TAG POS=$p TYPE=TD FORM=ID:frmASNCreate ATTR=CLASS:$s EXTRACT=TXT"
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "SAVEAS TYPE=EXTRACT FOLDER=$SaveDir FILE=AtechCheck-$PONumber.csv"
						if($s -eq "tblDataShadeRowBorder")
						{
							$p = $p + 8
						}
						if($s -eq "tblDataRegRowBorder")
						{
							$s = "tblDataShadeRowBorder"
						}
						elseif($s -eq "tblDataShadeRowBorder")
						{
							$s = "tblDataRegRowBorder"
						}
						$c++
						$PONumberLines = $Shipment[$c].PONUMBER
					}
					Add-Content -Encoding UTF8 $rootpath\AtechExhaustCheck.iim "SAVEAS TYPE=TXT FOLDER=$SaveDir FILE=done.txt"
					
					if ((Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe") -eq $true)
					{
						$cmdLine = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
					}else
					{
						$cmdLine = "C:\Program Files\Mozilla Firefox\firefox.exe"
					}
					$args = "imacros://run/?m=AtechExhaustCheck.iim"
					start-process $cmdLine $args
					Get-Process | ? {$_.Name -like "firefox"} | %{$_.Close()}
					
					
					while (!(Test-Path $SaveDir\done.txt))
					{
						#Wait Until exists!
					}
					Remove-Item $SaveDir\done.txt				
					$CheckPNX = Get-Content $SaveDir\AtechCheck-$PONumber.csv
					$Check = $CheckPNX | ForEach-Object {$_ -replace '"', ""} | ForEach-Object { $_.Remove(0,4) }
					if((AreArraysEqual $Check $ProdCheck) -eq $false)
					{
					
						#####################################################
						#Check failed, bail out!
						#####################################################
					
						$Errors = $Errors + "$PONumber has a mismatched line. Either an incorrect part number was entered or it was entered out of order."
                        $i = $v
					}	
					elseif((AreArraysEqual $Check $ProdCheck) -eq $true)
					{
						If((Test-Path $rootpath\AtechExhaustASN.iim) -eq $true)
						{
							Remove-Item $rootpath\AtechExhaustASN.iim
						}
						########################################################
						#Check is ok! Continue to create ASN
						########################################################
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Start at ASN Creation Screen. User must be logged in!
						`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Fill in PO#"
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Fill in the Quantity Shipped"
						$q = 1
					
						#################################################
						#Fill in Item Quantity							
						#################################################
						
						while ($Shipment[$i].PONUMBER -eq $PONumber)
						{	
	
							#################################################
							#Line Complete?									
							#################################################
	
							$LineComplete = $Shipment[$i].LINECOMPLETE
							if($LineComplete -eq "Y")
							{
								Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:frmASNCreate ATTR=ID:checkbox_$q CONTENT=YES"
							}
							else
							{
								$QTYShipped = $Shipment[$i].QTYSHIPPED
								Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=ID:ASNQty$q CONTENT=$QTYShipped"
							}
							
						
				
							$i++
							$q++
						}
						
						#################################################
						#Determine Carrier Code							
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
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label ")
						{
							$Carrier = "%0040"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label Saturday Delivery")
						{
							$Carrier = "%0043"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq '"Roadrunner"')
						{
							$Carrier = "%0034"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq "Roadrunner")
						{
							$Carrier = "%0034"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq '"UPS Freight LTL Standard"')
						{
							$Carrier = "%0055"
						}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Freight LTL Standard")
						{
						$Carrier = "%0055"
						}else
						{
							$Carrier = "%0023"
							$Comment = $Shipment[$i-1].SHIPPERNAME
							Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=TEXTAREA FORM=NAME:frmASNCreate ATTR=NAME:Comments CONTENT=`"$Comment`""
						}
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Choose Carrier `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:CarrierID CONTENT=$Carrier"
				
						$DateShipped = $Shipment[$i-1].SHIPDATE
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Shipment Date `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:Arrival CONTENT=%$DateShipped"
					
						$TrackingNumber = $Shipment[$i-1].TRACKINGNO
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Tracking Number Insertion`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:BOL CONTENT=$TrackingNumber"
				
						#########################################
						#Insert InvoiceNumber					
						#########################################
						
						$InvoiceNumber = $Shipment[$i-1].INVOICENUMBER
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`n'Add InvoiceNumber to Invoice Field`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:VendorReference CONTENT=$InvoiceNumber"
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "WAIT SECONDS = 5"
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:frmASNCreate ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmApprove ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
				
						#########################################
						#Save ASN as a Webpage					
						#########################################
						if((Test-Path $SaveDir\$Date\$PONumber-$Date) -eq $true)
						{
							Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date-extra"
						}
						else
						{
							Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date"
						}
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "SAVEAS TYPE=TXT FOLDER=$SaveDir\ FILE=done.txt"

					}
				}
				#########################################
				#Logout									
				#########################################
			
				if($i -eq $Shipment.length)
				{
						Add-Content -Encoding UTF8 $rootpath\AtechExhaustASN.iim "`nTAG POS=1 TYPE=A ATTR=TXT:Log<SP>Out"
				}				
				
				#####################################################
				#Run the Macro in Firefox, Firefox must be started!	
				#####################################################
			
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
				
				while (!(Test-Path $SaveDir\done.txt))
				{
					#Wait Until exists!
				}
				
				$f = Get-Content $SaveDir\done.txt | Select-String "No Results Found" -quiet
				if($f -eq $true)
				{
					$Errors = $Errors + "$PONumber is not a valid PO and had a fatal error"
					Remove-Item "$SaveDir\$Date\$PONumber-$Date"
				}
				$f = $null
				$f = Get-Content $SaveDir\done.txt | Select-String "No items found with search criteria." -quiet
				if($f -eq $true)
				{
					$Errors = $Errors + "$PONumber was probably already finished. If not please run manually."
				}
					
				Remove-Item $SaveDir\done.txt
	
			
				
				$LineComplete = $null
				$ProdCheck = $null
			
			}
				
				

			#########################################
			#End Loop Here							
			#########################################
			if($Errors.length -gt 0)
			{
				[reflection.assembly]::loadwithpartialname('system.windows.forms');
				[system.Windows.Forms.MessageBox]::show($Errors)
				[reflection.assembly]::loadwithpartialname('system.windows.forms');
				[system.Windows.Forms.MessageBox]::show("If you did not catch that, the errors are logged in the ASN folder on the Snap Drive under todays date.")
				
				if($Errors.length -gt 0)
				{
					$Errors | Export-CSV $SaveDir\$Date\Exhaust-Errors.csv
				}
				$Errors = $Null
			}
			$Continue = [System.Windows.Forms.MessageBox]::Show("ASN creation for Exhaust should now be complete. Would you like to run another division?" , "Status" , 4)
		}
		
		Else
		{
			$b = [System.Windows.Forms.MessageBox]::Show("Please make sure you downloaded the Exhaust CSV from Pentaho. Would you like to start over?" , "Status" , 4)
			if($b -eq "No")
			{
				exit
			}
		}
        
        $i = $null
		$b = $null
		$a = $null
		$v = $null
		$Division = $null
    }

	$Shipment = $null
	#####################################################
	#Start Private Label block of code						
	#####################################################
	
	if ($Division -eq "PrivateLabel")
	{
		#########################################################
		#Create Error Array
		#########################################################
		
		$Errors = @()

		if ((Test-Path $DownloadDir\AtechExhPrivateLabelASN.csv) -eq $true)
		{
		
			#############################################
			#Remove old files							
			#############################################
			
			if ((Test-Path $rootpath\AtechExhPrivateLabelASN.iim) -eq $true)
			{    
				Remove-Item $rootpath\AtechExhPrivateLabelASN.iim
			}
			if ((Test-Path $rootpath\AtechExhPrivateLabelLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechExhPrivateLabelLogin.iim
			}
			
			#############################################
			#Login information							
			#############################################

			$username = Read-Host 'Please input your Username for Private Label'
			$pass = Read-Host 'Please input your password?' -AsSecureString
			
			#############################################
			#Convert password to plain text variable	
			#############################################
			
			$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))
			
			#############################################
			#Build Login Macro 							
			#############################################
			
			Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelLogin.iim "TAB T=1 `nURL GOTO=https://b2b.atechmotorsports.com/default.asp"
			Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelLogin.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmLogin ATTR=NAME:UserName CONTENT=$username"
			Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelLogin.iim "SET !ENCRYPTION NO"
			Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelLogin.iim "TAG POS=1 TYPE=INPUT:PASSWORD FORM=NAME:frmLogin ATTR=NAME:Password CONTENT=$password"
			Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelLogin.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmLogin ATTR=VALUE:Go"
			Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelLogin.iim "SAVEAS TYPE=TXT FOLDER=$CSVDir FILE=login.txt"
			$password = $null
			
			
			#############################################
			#Run Login Macro							
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
    			
			while (!(Test-Path $CSVDir\login.txt))
			{
				#Wait Until exists!
			}
			
			
			$b = [System.Windows.Forms.MessageBox]::Show("Did we login to Atech correctly?" , "Status" , 4)
			if($b -eq "No")
			{
				exit
			}
			
			if ((Test-Path $CSVDir\login.txt) -eq $true)
			{
				Remove-Item $CSVDir\login.txt
			}
			
			if ((Test-Path $rootpath\AtechExhPrivateLabelLogin.iim) -eq $true)
			{
				Remove-Item $rootpath\AtechExhPrivateLabelLogin.iim
			}
			
			#############################################
			#Move CSV to correct Location				
			#############################################
	
			Move-Item -path $DownloadDir\AtechExhPrivateLabelASN.csv -destination $SaveDir\CSV\$Date-ExhPrivateLabel.csv
			$Shipment = Import-CSV "$SaveDir\CSV\$Date-ExhPrivateLabel.csv"
			if ((Test-Path $DownloadDir\AtechExhPrivateLabelASN.csv) -eq $true)
			{
				[System.Windows.Forms.MessageBox]::Show("After this point you will need to download a new CSV if there are any errors") | Out-Null
				Remove-Item $DownloadDir\AtechExhPrivateLabelASN.csv
			}
			
			#####################################################
			#Determine if this is for the first line or not.	
			#####################################################
			
			if ($i -lt 1)
			{
				$i = 0
			}
			
			#########################################################
			#Loop Beginning so that script knows where to stop from
			#CSV Length
			#########################################################
			
			
			while ($i -le $Shipment.length-1)
			{			

				#####################################################
				#Begin Checking
				#####################################################
				
				$c = $i
				$v = $i
				$PONumber = $Shipment[$i].PONUMBER
				$POLineComplete = $PONumber
				$LineComplete = @()
				$ProdCheck = @()
				
				###################################################################
				#Build 2 Arrays. One for Checking if the PO is shipping complete
				#The other is to store the product numbers if it is not complete
				###################################################################
				
				while($POLineComplete -eq $PONumber)
				{
					$LineComplete = $LineComplete + $Shipment[$c].LINECOMPLETE
					$ProdCheck = $ProdCheck + $Shipment[$c].PRODNUM
					$c++
					$v++
					$POLineComplete = $Shipment[$v].PONUMBER
				}
				
				if($LineComplete -notcontains "N")
				{
					if((Test-Path $rootpath\AtechExhPrivateLabelASN.iim) -eq $true)
					{
						Remove-Item $rootpath\AtechExhPrivateLabelASN.iim
					}
					##############################################
					#Ship complete Piece of CAKE!
					##############################################
					
					#####################################################
					#Insert PO and Head to Creation Form				
					#####################################################
				
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "TAB T=1"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Start at ASN Creation Screen.
					`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "SET !ERRORIGNORE YES"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "SET !TIMEOUT_STEP 0"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Fill in PO#"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Fill in the Quantity Shipped"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:frmASNCreate ATTR=NAME:CheckAllBoxes CONTENT=YES"
				
				
					#################################################
					#Determine Carrier Code							
					#################################################
					
					if ($Shipment[$i].SHIPPERNAME -eq "UPS  - Ground")
					{
						$Carrier = "%0001"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Residential ")
					{    
						$Carrier = "%0001"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "FedEx  - Ground")
					{
						$Carrier = "%0005"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Red Label Res'd")
					{
						$Carrier = "%0040"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Blue Label Res'd")
					{
						$Carrier = "%0041"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Blue Label ")
					{
						$Carrier = "%0041"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Red Label ")
					{
						$Carrier = "%0040"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Red Label Saturday Delivery")
					{
						$Carrier = "%0043"
					}elseif ($Shipment[$i].SHIPPERNAME -eq '"Roadrunner"')
					{
						$Carrier = "%0034"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "Roadrunner")
					{
						$Carrier = "%0034"
					}elseif ($Shipment[$i].SHIPPERNAME -eq '"UPS Freight LTL Standard"')
					{
						$Carrier = "%0055"
					}elseif ($Shipment[$i].SHIPPERNAME -eq "UPS Freight LTL Standard")
					{
						$Carrier = "%0055"
					}else
					{
						$Carrier = "%0023"
						$Comment = $Shipment[$i].SHIPPERNAME
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=TEXTAREA FORM=NAME:frmASNCreate ATTR=NAME:Comments CONTENT=`"$Comment`""
					}
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Choose Carrier `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:CarrierID CONTENT=$Carrier"
			
					$DateShipped = $Shipment[$i].SHIPDATE
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Shipment Date `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:Arrival CONTENT=%$DateShipped"
				
					$TrackingNumber = $Shipment[$i].TRACKINGNO
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Tracking Number Insertion`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:BOL CONTENT=$TrackingNumber"
			
					#########################################
					#Insert InvoiceNumber					
					#########################################
					
					$InvoiceNumber = $Shipment[$i].INVOICENUMBER
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Add InvoiceNumber to Invoice Field`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:VendorReference CONTENT=$InvoiceNumber"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:frmASNCreate ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmApprove ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
			
					#########################################
					#Save ASN as a Webpage					
					#########################################
					if((Test-Path $SaveDir\$Date\$PONumber-$Date.htm) -eq $true)
					{
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date-extra"
					}
					else
					{
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date"
					}
					
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "SAVEAS TYPE=TXT FOLDER=$SaveDir\ FILE=done.txt"
					
					$i = $v
					
				}
				
				else
				{
					##############################################
					#Continue checking!
					##############################################
					$c = $i
					$s = "tblDataRegRowBorder"
                    $p = 3
					$PONumberLines = $Shipment[$i].PONUMBER
					if((Test-Path $rootpath\AtechExhPrivateLabelCheck.iim) -eq $true)
					{
						Remove-Item $rootpath\AtechExhPrivateLabelCheck.iim
					}
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "TAB T=1"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "`n'Start at ASN Creation Screen.
					`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "SET !EXTRACT_TEST_POPUP NO"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "SET !ERRORIGNORE YES"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "SET !TIMEOUT_STEP 0"
					#New Stuff
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "`n'Fill in PO#"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
			
                    while($PONumberLines -eq $PONumber)
					{
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "TAG POS=$p TYPE=TD FORM=ID:frmASNCreate ATTR=CLASS:$s EXTRACT=TXT"
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "SAVEAS TYPE=EXTRACT FOLDER=$SaveDir FILE=AtechCheck-$PONumber.csv"
						if($s -eq "tblDataShadeRowBorder")
						{
							$p = $p + 8
						}
						if($s -eq "tblDataRegRowBorder")
						{
							$s = "tblDataShadeRowBorder"
						}
						elseif($s -eq "tblDataShadeRowBorder")
						{
							$s = "tblDataRegRowBorder"
						}
						$c++
						$PONumberLines = $Shipment[$c].PONUMBER
					}
					Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelCheck.iim "SAVEAS TYPE=TXT FOLDER=$SaveDir FILE=done.txt"
					
					if ((Test-Path "C:\Program Files (x86)\Mozilla Firefox\firefox.exe") -eq $true)
					{
						$cmdLine = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
					}else
					{
						$cmdLine = "C:\Program Files\Mozilla Firefox\firefox.exe"
					}
					$args = "imacros://run/?m=AtechExhPrivateLabelCheck.iim"
					start-process $cmdLine $args
					Get-Process | ? {$_.Name -like "firefox"} | %{$_.Close()}
					
					
					while (!(Test-Path $SaveDir\done.txt))
					{
						#Wait Until exists!
					}
					Remove-Item $SaveDir\done.txt				
					$CheckPNX = Get-Content $SaveDir\AtechCheck-$PONumber.csv
					$Check = $CheckPNX | ForEach-Object {$_ -replace '"', ""} | ForEach-Object { $_.Remove(0,4) }
					if((AreArraysEqual $Check $ProdCheck) -eq $false)
					{
					
						#####################################################
						#Check failed, bail out!
						#####################################################
					
						$Errors = $Errors + "$PONumber has a mismatched line. Either an incorrect part number was entered or it was entered out of order."
                        $i = $v
					}	
					elseif((AreArraysEqual $Check $ProdCheck) -eq $true)
					{
						If((Test-Path $rootpath\AtechExhPrivateLabelASN.iim) -eq $true)
						{
							Remove-Item $rootpath\AtechExhPrivateLabelASN.iim
						}
						########################################################
						#Check is ok! Continue to create ASN
						########################################################
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Start at ASN Creation Screen. User must be logged in!
						`nURL GOTO=https://b2b.atechmotorsports.com/ASNCreate.asp?Function=POForm"
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Fill in PO#"
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmKnownPO ATTR=NAME:PONumber CONTENT=$PONumber"
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmKnownPO ATTR=VALUE:Search<SP>for<SP>P.O."
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Fill in the Quantity Shipped"
						$q = 1
					
						#################################################
						#Fill in Item Quantity							
						#################################################
						
						while ($Shipment[$i].PONUMBER -eq $PONumber)
						{	
	
							#################################################
							#Line Complete?									
							#################################################
	
							$LineComplete = $Shipment[$i].LINECOMPLETE
							if($LineComplete -eq "Y")
							{
								Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:CHECKBOX FORM=NAME:frmASNCreate ATTR=ID:checkbox_$q CONTENT=YES"
							}
							else
							{
								$QTYShipped = $Shipment[$i].QTYSHIPPED
								Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "TAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=ID:ASNQty$q CONTENT=$QTYShipped"
							}
							
						
				
							$i++
							$q++
						}
						
						#################################################
						#Determine Carrier Code							
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
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label ")
						{
							$Carrier = "%0040"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Red Label Saturday Delivery")
						{
							$Carrier = "%0043"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq '"Roadrunner"')
						{
							$Carrier = "%0034"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq "Roadrunner")
						{
							$Carrier = "%0034"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq '"UPS Freight LTL Standard"')
						{
							$Carrier = "%0055"
						}elseif ($Shipment[$i-1].SHIPPERNAME -eq "UPS Freight LTL Standard")
						{
							$Carrier = "%0055"
						}else
						{
							$Carrier = "%0023"
							$Comment = $Shipment[$i-1].SHIPPERNAME
							Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=TEXTAREA FORM=NAME:frmASNCreate ATTR=NAME:Comments CONTENT=`"$Comment`""
						}
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Choose Carrier `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:CarrierID CONTENT=$Carrier"
				
						$DateShipped = $Shipment[$i-1].SHIPDATE
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Shipment Date `nTAG POS=1 TYPE=SELECT FORM=NAME:frmASNCreate ATTR=NAME:Arrival CONTENT=%$DateShipped"
					
						$TrackingNumber = $Shipment[$i-1].TRACKINGNO
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Tracking Number Insertion`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:BOL CONTENT=$TrackingNumber"
				
						#########################################
						#Insert InvoiceNumber					
						#########################################
						
						$InvoiceNumber = $Shipment[$i-1].INVOICENUMBER
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`n'Add InvoiceNumber to Invoice Field`nTAG POS=1 TYPE=INPUT:TEXT FORM=NAME:frmASNCreate ATTR=NAME:VendorReference CONTENT=$InvoiceNumber"
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "WAIT SECONDS = 5"
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=ID:frmASNCreate ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=INPUT:SUBMIT FORM=NAME:frmApprove ATTR=VALUE:<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>Accept<SP><SP><SP><SP><SP><SP><SP><SP><SP><SP>"
				
						#########################################
						#Save ASN as a Webpage					
						#########################################
						if((Test-Path $SaveDir\$Date\$PONumber-$Date) -eq $true)
						{
							Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date-extra"
						}
						else
						{
							Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nSAVEAS TYPE=HTM FOLDER=$SaveDir\$Date\ FILE=$PONumber-$Date"
						}
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "SAVEAS TYPE=TXT FOLDER=$SaveDir\ FILE=done.txt"

					}
				}
				#########################################
				#Logout									
				#########################################
			
				if($i -eq $Shipment.length)
				{
						Add-Content -Encoding UTF8 $rootpath\AtechExhPrivateLabelASN.iim "`nTAG POS=1 TYPE=A ATTR=TXT:Log<SP>Out"
				}				
				
				#####################################################
				#Run the Macro in Firefox, Firefox must be started!	
				#####################################################
			
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
				
				while (!(Test-Path $SaveDir\done.txt))
				{
					#Wait Until exists!
				}
				
				$f = Get-Content $SaveDir\done.txt | Select-String "No Results Found" -quiet
				if($f -eq $true)
				{
					$Errors = $Errors + "$PONumber is not a valid PO and had a fatal error"
					Remove-Item "$SaveDir\$Date\$PONumber-$Date"
				}
				$f = $null
				$f = Get-Content $SaveDir\done.txt | Select-String "No items found with search criteria." -quiet
				if($f -eq $true)
				{
					$Errors = $Errors + "$PONumber was probably already finished. If not please run manually."
				}
					
				Remove-Item $SaveDir\done.txt
	
			
				
				$LineComplete = $null
				$ProdCheck = $null
			
			}
				
				

			#########################################
			#End Loop Here							
			#########################################
			if($Errors.length -gt 0)
			{
				[reflection.assembly]::loadwithpartialname('system.windows.forms');
				[system.Windows.Forms.MessageBox]::show($Errors)
				[reflection.assembly]::loadwithpartialname('system.windows.forms');
				[system.Windows.Forms.MessageBox]::show("If you did not catch that, the errors are logged in the ASN folder on the Snap Drive under todays date.")
				
				if($Errors.length -gt 0)
				{
					$Errors | Export-CSV $SaveDir\$Date\ExhPrivateLabel-Errors.csv
				}
				$Errors = $Null
			}
			$Continue = [System.Windows.Forms.MessageBox]::Show("ASN creation for Private Label should now be complete. Would you like to run another division?" , "Status" , 4)
		}
		
		Else
		{
			$b = [System.Windows.Forms.MessageBox]::Show("Please make sure you downloaded the Private Label CSV from Pentaho. Would you like to start over?" , "Status" , 4)
			if($b -eq "No")
			{
				exit
			}
		}
        
        $i = $null
		$b = $null
		$a = $null
		$v = $null
		$Division = $null
    }

	$Shipment = $null

}