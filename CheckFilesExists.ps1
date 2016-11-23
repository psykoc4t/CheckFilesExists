<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.128
	 Created on:   	18/11/2016 17:00
	 Created by:   	Antonio de Almeida
	 Organization: 	antoniodealmeida.net
	 Version:		1.0
	 Filename:     	CheckFilesExists.ps1
	===========================================================================
	.DESCRIPTION
	Check files presence in folders.
#>

$scriptName = $MyInvocation.MyCommand.Name

if ($args.Count -lt 2)
		{
			Get-Date -uformat "%Hh%M(%S) : ERR. : Argument manquant'"
			 "waiting for: $scriptName path fileExtension"
			exit 10
		}

$NotifEmail = "adealmeida@netc.fr"

$dir = $args[0]
$extension = $args[1]

$formatedDir = $dir.Replace(':', "$")

$a = "<style>"
$a = $a + "BODY{background-color:white;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}"
$a = $a + "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:palegoldenrod}}"
$a = $a + "</style>"

if($extension -eq "ALL"){
"Extension ALL"

$check = @(gci -Path $dir -Force -File | select-object LastWriteTime, Name, @{Name="Kbytes";Expression={[math]::Round(($_.Length / 1Kb),2)}}| ConvertTo-Html -head $a)
$entries = gci -Path $dir -Force -File
$count = $entries.count
}else{
"Extension .$extension"
$check = @(gci -Path $dir -Force -File | where-object {$_.extension -eq ".$extension"} | select-object lastWriteTime, name, @{Name="Kbytes"; Expression={[math]::Round(($_.Length / 1Kb),2)}} | ConvertTo-Html -head $a)
$entries = gci -Path $dir -Force -File | where-object {$_.extension -eq ".$extension"} 
$count = $entries.count
}






###### DEBUT DE LA FONCTIONS MAIL ######
Function SendMail
{
	
	param (
		
		[Parameter(Mandatory = $False)]
		[string]$from = "$env:computername@antoniodealmeida.net",
		[Parameter(Mandatory = $false)]
		[string]$to = $NotifEmail,

		[Parameter(Mandatory = $false)]
		[string]$cc = "",

		[Parameter(Mandatory = $false)]
		[string]$Subject = "Check file ($count) $dir",
		[Parameter(Mandatory = $false)]
		$body,

		[Parameter(Mandatory = $false)]
		[string]$pj,
		[Parameter(Mandatory = $false)]
		[string]$smtpserver = "serveur_smtp"
		
	)
	
	# TEMPLATE DU MESSAGE MAIL
	$body_html = @"
<br>
Bonjour,
<br><br>
Le dossier \\$env:computername\$formatedDir contient:
<br><br>
$body
<br><br>
******************************
$PSScriptRoot\$scriptName
******************************

"@
	
	# CONSTRUCTION DU MAIL
	$message = new-object System.Net.Mail.MailMessage
	
	# FROM
	$message.From = $from
	# TO
	if ("$to" -eq "") { $message.To.Add("adealmeida@netc.fr") }
	else { $message.To.Add($to) }
	# CC
	if ("$cc" -ne "") { $message.To.Add($cc) }
	
	# ENVOI AU FORMAT HTML
	$message.IsBodyHtml = $True
	
	# PIECE JOINTES
	if ("$pj" -ne "")
	{
		$attach = new-object Net.Mail.Attachment($pj)
		$message.Attachments.Add($attach)
	}
	
	# SUJET
	$message.Subject = $Subject
	$message.body = $body_html
	
	
	# ENVOI DU MAIL
	$smtp = new-object Net.Mail.SmtpClient($smtpserver)
	$smtp.Send($message)
	if ($? -eq $true) { Get-Date -uformat "%Hh%M(%S) : MAIL ENVOYE" }
	
	
}
###### FIN DE LA FONCTIONS MAIL ######
if($check.length -gt 0){
"$count fichier(s) present(s)"
SendMail -body $check
}else{"Aucun fichier present dans $dir"}
