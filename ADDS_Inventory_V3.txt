#Requires -Version 3.0
#requires -Module ActiveDirectory
#requires -Module GroupPolicy
#This File is in Unicode format.  Do not edit in an ASCII editor. Notepad++ UTF-8-BOM

<#
.SYNOPSIS
	Creates a complete inventory of a Microsoft Active Directory Forest or Domain.
.DESCRIPTION
	Creates a complete inventory of a Microsoft Active Directory Forest or Domain using 
	Microsoft PowerShell, Word, plain text, or HTML.
	
	Creates a Word or PDF document, text, or HTML file named after the Active Directory 
	Forest or Domain.
	
	Version 3.0 changes the default output report from Word to HTML.
	
	Word and PDF document includes a Cover Page, Table of Contents, and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

	The script requires at least PowerShell version 3 but runs best in version 5.

	Word is NOT needed to run the script. This script outputs in Text and HTML.
	
	You do NOT have to run this script on a domain controller. This script was developed 
	and run from a Windows 10 VM.

	While most of the script can run with a non-admin account, there are some features 
	that will not or may not work without domain admin or enterprise admin rights.  
	The Hardware and Services parameters require domain admin privileges.  
	
	The script does gathering of information on Time Server and AD database, log file, and 
	SYSVOL locations. Those require access to the registry on each domain controller, which 
	means the script should now always be run from an elevated PowerShell session with an 
	account with a minimum of domain admin rights.
	
	Running the script in a forest with multiple domains requires Enterprise Admin rights.

	The count of all users may not be accurate if the user running the script does not have 
	the necessary permissions on all user objects.  In that case, there may be user accounts 
	classified as "unknown".
	
	To run the script from a workstation, RSAT is required.
	
	Remote Server Administration Tools for Windows 7 with Service Pack 1 (SP1)
		https://carlwebster.sharefile.com/d-sace5ee0f1ada47289ca14be544878a24
		
	Remote Server Administration Tools for Windows 8 
		https://carlwebster.sharefile.com/d-s791075d451fc415ca83ec8958b385a95
		
	Remote Server Administration Tools for Windows 8.1 
		https://carlwebster.sharefile.com/d-s37209afb73dc48f497745924ed854226
		
	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520
	
.PARAMETER ADDomain
	Specifies an Active Directory domain object by providing one of the following 
	property values. The identifier in parentheses is the LDAP display name for the 
	attribute. All values are for the domainDNS object that represents the domain.

	Distinguished Name

	Example: DC=labaddomain,DC=com

	GUID (objectGUID)

	Example: b147a813-9938-4a89-bc93-58a0d36c41c3

	Security Identifier (objectSid)

	Example: S-1-5-21-3916992870-515249095-1421388220

	DNS domain name

	Example: labaddomain.com

	NetBIOS domain name

	Example: labaddomain

	If both ADForest and ADDomain are specified, ADDomain takes precedence.
.PARAMETER ADForest
	Specifies an Active Directory forest object by providing one of the following 
	attribute values. 
	The identifier in parentheses is the LDAP display name for the attribute.

	Fully qualified domain name
		Example: labaddomain.com
	GUID (objectGUID)
		Example: b147a813-9938-4a89-bc93-58a0d36c41c3
	DNS host name
		Example: labaddomain.com
	NetBIOS name
		Example: labaddomain
	
	Default value is $Env:USERDNSDOMAIN	
	
	If both ADForest and ADDomain are specified, ADDomain takes precedence.
.PARAMETER ComputerName
	Specifies which domain controller to use to run the script against.
	If ADForest is a trusted forest, then ComputerName is required to detect the 
	existence of ADForest.
	ComputerName can be entered as the NetBIOS name, FQDN, localhost, or IP Address.
	If entered as localhost, the actual computer name is determined and used.
	If entered as an IP address, an attempt is made to determine and use the actual 
	computer name.
	
	This parameter has an alias of ServerName.
	Default value is $Env:USERDNSDOMAIN	
.PARAMETER MaxDetails
	Adds maximum detail to the report.
	
	This is the same as using the following parameters:
		DCDNSInfo
		GPOInheritance
		Hardware
		IncludeUserInfo
		Services
	
	WARNING: Using this parameter can create an extremely large report and 
	can take a very long time to run.

	This parameter has an alias of MAX.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	
	HTML is now the default report format.
	
	This parameter is set True if no other output format is selected.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	
	This parameter is disabled by default.
.PARAMETER AddDateTime
	Adds a date timestamp to the end of the file name.
	
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2022 at 6PM is 2022-06-01_1800.
	
	Output filename will be ReportName_2022-06-01_1800.docx (or .pdf).
	
	This parameter is disabled by default.
	This parameter has an alias of ADT.
.PARAMETER CompanyName
	Company Name to use for the Word Cover Page or the Forest Information section for 
	HTML and Text.
	
	Default value for Word output is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.

	This parameter has an alias of CN.

	For Word output, if either registry key does not exist and this parameter is not 
	specified, the report does not contain a Company Name on the cover page.
	
	For HTML and Text output, the Forest Information section does not contain the Company 
	Name if this parameter is not specified.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script runs.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script runs.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER ReportFooter
	Outputs a footer section at the end of the report.

	This parameter has an alias of RF.
	
	Report Footer
		Report information:
			Created with: <Script Name> - Release Date: <Script Release Date>
			Script version: <Script Version>
			Started on <Date Time in Local Format>
			Elapsed time: nn days, nn hours, nn minutes, nn.nn seconds
			Ran from domain <Domain Name> by user <Username>
			Ran from the folder <Folder Name>

	Script Name and Script Release date are script-specific variables.
	Start Date Time in Local Format is a script variable.
	Elapsed time is a calculated value.
	Domain Name is $env:USERDNSDOMAIN.
	Username is $env:USERNAME.
	Folder Name is a script variable.
.PARAMETER DCDNSInfo 
	Use WMI to gather, for each domain controller, the IP Address, and each DNS server 
	configured.
	This parameter requires the script to run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e., Domain 
	Admin).
	Selecting this parameter adds an extra section to the report.
	This parameter is disabled by default.
.PARAMETER GPOInheritance
	In the Group Policies by OU section adds Inherited GPOs in addition to the GPOs 
	directly linked.
	Adds a second column to the table GPO Type.
	
	This parameter is disabled by default.
	This parameter has an alias of GPO.
.PARAMETER Hardware
	Use WMI to gather hardware information on Computer System, Disks, Processor, and 
	Network Interface Cards
	
	This parameter requires the script to run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e., Domain 
	Admin).
	
	Selecting this parameter will add to both the time it takes to run the script and 
	size of the report.
	
	This parameter is disabled by default.
.PARAMETER IncludeUserInfo
	For the User Miscellaneous Data section outputs a table with the SamAccountName
	and DistinguishedName of the users in the All Users counts:
	
		Disabled users
		Unknown users
		Locked out users
		All users with password expired
		All users whose password never expires
		All users with password not required
		All users who cannot change password
		All users with SID History
		All users with Homedrive set in ADUC
		All users whose Primary Group is not Domain Users
		All users with RDS HomeDrive set in ADUC
		All Names in the ForeignSecurityPrincipals container that are orphan SIDs 
		(Root domain only)
	
	The Text output option is limited to the first 25 characters of the SamAccountName
	and the first 116 characters of the DistinguishedName.
	
	This parameter is disabled by default.
	This parameter has an alias of IU.
.PARAMETER Section
	Processes one or more sections of the report.
	Valid options are:
		Forest
		Sites
		Domains (includes Domain Controllers and optional Hardware, Services and 
		DCDNSInfo)
		OUs (Organizational Units)
		Groups
		GPOs
		Misc (Miscellaneous data)
		All

	This parameter defaults to All sections.
	
	Multiple sections are separated by a comma. -Section forest, domains
.PARAMETER Services
	Gather information on all services running on domain controllers.
	
	Services that are configured to automatically start but are not running will be 
	colored in red.
	
	This parameter requires the script be run from an elevated PowerShell session
	using an account with permission to retrieve service information (i.e. Domain 
	Admin).
	
	Selecting this parameter will add to both the time it takes to run the script and 
	size of the report.
	
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	
	Microsoft Word is no longer the default report format.
	This parameter is disabled by default.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	
	This parameter requires Microsoft Word to be installed.
	This parameter uses Word's SaveAs PDF capability.

	This parameter is disabled by default.
.PARAMETER CompanyAddress
	Company Address to use for the Cover Page, if the Cover Page has the Address field.
	
	The following Cover Pages have an Address field:
		Banded (Word 2013/2016)
		Contrast (Word 2010)
		Exposure (Word 2010)
		Filigree (Word 2013/2016)
		Ion (Dark) (Word 2013/2016)
		Retrospect (Word 2013/2016)
		Semaphore (Word 2013/2016)
		Tiles (Word 2010)
		ViewMaster (Word 2013/2016)
		
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email to use for the Cover Page, if the Cover Page has the Email field.  
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover Page, if the Cover Page has the Fax field.  
	
	The following Cover Pages have a Fax field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyPhone
	Company Phone to use for the Cover Page if the Cover Page has the Phone field.  
	
	The following Cover Pages have a Phone field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013, and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly 
		works in 2010 but Subtitle/Subject & Author fields need to be moved 
		after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 
		36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually 
		resized or font changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 
		2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	The default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER SmtpPort
	Specifies the SMTP port for the SmtpServer. 

	The default is 25.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report(s). 
	
	If From or To are used, this is a required parameter.
.PARAMETER From
	Specifies the username for the From email address.
	
	If SmtpServer or To are used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	
	If SmtpServer or From are used, this is a required parameter.
.PARAMETER UserName
	Username to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	The default is False.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1
	
	Creates an HTML report.
	
	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -MSWord
	
	Uses all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -ADForest company.tld
	
	Creates an HTML report.
	
	company.tld for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -ADDomain child.company.tld
	
	Creates an HTML report.
	
	child.company.tld for the AD Domain.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -ADForest parent.company.tld -ADDomain 
	child.company.tld
	
	Creates an HTML report.
	
	Because both ADForest and ADDomain are specified, ADDomain wins and child.company.tld 
	is used for AD Domain.
	
	ADForest is set to the value of ADDomain.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -ADForest company.tld -ComputerName DC01 
	-MSWord
	
	Creates a Microsoft Word report.
	
	Uses all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	company.tld for the AD Forest
	
	The script runs remotely on the DC01 domain controller.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -PDF -ADForest corp.carlwebster.com
	
	Uses all default values and saves the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	corp.carlwebster.com for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -Text -ADForest corp.carlwebster.com
	
	Uses all default values and saves the document as a formatted text file.

	corp.carlwebster.com for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -HTML -ADForest corp.carlwebster.com
	
	Uses all default values and saves the document as an HTML file.

	corp.carlwebster.com for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -hardware
	
	Creates an HTML report.
	
	Uses all default values and adds additional information for each domain controller about 
	its hardware.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -services
	
	Creates an HTML report.
	
	Will use all default values and add additional information for the services running 
	on each domain controller.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -DCDNSInfo
	
	Creates an HTML report.
	
	Uses all default values and adds additional information for each domain controller about 
	its DNS IP configuration.

	An extra section is added to the end of the report.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory_V3.ps1 -MSWord -CompanyName "Carl Webster 
	Consulting" -CoverPage "Mod" -UserName "Carl Webster" -ComputerName ADDC01

	Creates a Microsoft Word report.
	
	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	Domain Controller named ADDC01 for the ComputerName.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory_V3.ps1 -MSWord -CN "Carl Webster Consulting" 
	-CP "Mod" -UN "Carl Webster"

	Creates a Microsoft Word report.
	
	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory_V3.ps1 -MSWord -CompanyName "Sherlock Holmes 
    Consulting" -CoverPage Exposure -UserName "Dr. Watson" -CompanyAddress "221B Baker 
    Street, London, England" -CompanyFax "+44 1753 276600" -CompanyPhone "+44 1753 276200
	
	Creates a Microsoft Word report.
	
	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Exposure for the Cover Page format.
		Dr. Watson for the User Name.
		221B Baker Street, London, England for the Company Address.
		+44 1753 276600 for the Company Fax.
		+44 1753 276200 for the Company Phone.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory_V3.ps1 -MSWord -CompanyName "Sherlock Holmes 
	Consulting" -CoverPage Facet -UserName "Dr. Watson" -CompanyEmail 
	SuperSleuth@SherlockHolmes.com

	Creates a Microsoft Word report.
	
	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Facet for the Cover Page format.
		Dr. Watson for the User Name.
		SuperSleuth@SherlockHolmes.com for the Company Email.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -ADForest company.tld -AddDateTime
	
	Creates an HTML report.

	company.tld for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.

	Adds a date time stamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2022 at 6PM is 2022-06-01_1800.
	The output filename is company.tld_2022-06-01_1800.docx.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -PDF -ADForest corp.carlwebster.com 
	-AddDateTime
	
	Uses all default values and saves the document as a PDF file.

	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	corp.carlwebster.com for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.

	Adds a date time stamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2022 at 6PM is 2022-06-01_1800.
	The output filename is corp.carlwebster.com_2022-06-01_1800.PDF
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -ADForest corp.carlwebster.com -Folder 
	\\FileServer\ShareName
	
	Creates an HTML report.

	corp.carlwebster.com for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
	
	The output file is saved in the path \\FileServer\ShareName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -Section Forest

	Creates an HTML report.

	ADForest defaults to the value of $Env:USERDNSDOMAIN

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and uses that as the value 
	for ComputerName.
	
	The report includes only the Forest section.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -Section groups, misc -ADForest 
	WebstersLab.com -ServerName PrimaryDC.websterslab.com

	Creates an HTML report.

	WebstersLab.com for ADForest.
	PrimaryDC.websterslab.com for ComputerName.
	
	The report includes only the Groups and Miscellaneous sections.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -MaxDetails
	
	Creates an HTML report.

	Set the following parameter values:
		DCDNSInfo       = True
		GPOInheritance  = True
		Hardware        = True
		IncludeUserInfo = True
		Services        = True
		
		Section         = "All"
.EXAMPLE
	PS C:\PSScript >.\ADDS_Inventory_V3.ps1 -Dev -ScriptInfo -Log
	
	Creates a default report.
	
	Creates a text file named ADDSInventoryScriptErrors_yyyyMMddTHHmmssffff.txt that 
	contains up to the last 250 errors reported by the script.
	
	Creates a text file named ADDSInventoryScriptInfo_yyyy-MM-dd_HHmm.txt that 
	contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	ADDSDocScriptTranscript_yyyyMMddTHHmmssffff.txt.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -SmtpServer mail.domain.tld -From 
	ADAdmin@domain.tld -To ITGroup@domain.tld	

	The script uses the email server mail.domain.tld, sending from ADAdmin@domain.tld 
	and sending to ITGroup@domain.tld.

	The script uses the default SMTP port 25 and does not use SSL.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -SmtpServer mailrelay.domain.tld -From 
	Anonymous@domain.tld -To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script uses the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld and sending to ITGroup@domain.tld.

	To send an unauthenticated email using an email relay server requires the From email 
	account to use the name Anonymous.

	The script uses the default SMTP port 25 and does not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send an email using a Gmail or g-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script generates an anonymous, secure password for the anonymous@domain.tld 
	account.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -SmtpServer 
	labaddomain-com.mail.protection.outlook.com -UseSSL -From 
	SomeEmailAddress@labaddomain.com -To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multiFunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script uses the email server labaddomain-com.mail.protection.outlook.com, sending 
	from SomeEmailAddress@labaddomain.com and sending to ITGroupDL@labaddomain.com.

	The script uses the default SMTP port 25 and SSL.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	The script uses the email server smtp.office365.com on port 587 using SSL, sending from 
	webster@carlwebster.com and sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V3.ps1 -SmtpServer smtp.gmail.com -SmtpPort 587 
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send an email using a Gmail or g-suite account, you may have to turn ON the "Less 
	secure app access" option on your account.
	*** NOTE ***
	
	The script uses the email server smtp.gmail.com on port 587 using SSL, sending from 
	webster@gmail.com and sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, the script prompts 
	the user to enter valid credentials.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word or PDF document.
.NOTES
	NAME: ADDS_Inventory_V3.ps1
	VERSION: 3.11
	AUTHOR: Carl Webster and Michael B. Smith
	LASTEDIT: May 27, 2022
#>


#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "") ]

Param(
	[parameter(Mandatory=$False)] 
	[string]$ADDomain="", 

	[parameter(Mandatory=$False)] 
	[string]$ADForest=$Env:USERDNSDOMAIN, 

	[parameter(Mandatory=$False)] 
	[Alias("ServerName")]
	[string]$ComputerName=$Env:USERDNSDOMAIN,
	
	[parameter(Mandatory=$False)] 
	[Alias("MAX")]
	[Switch]$MaxDetails=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(Mandatory=$False)] 
	[Alias("ADT")]
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,

	[parameter(Mandatory=$False)] 
	[Alias("RF")]
	[Switch]$ReportFooter=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$DCDNSInfo=$False, 

	[parameter(Mandatory=$False)] 
	[Switch]$GPOInheritance=$False, 

	[parameter(Mandatory=$False)] 
	[Switch]$Hardware=$False, 

	[parameter(Mandatory=$False)] 
	[Alias("IU")]
	[Switch]$IncludeUserInfo=$False,
	
	[Parameter( Mandatory = $False )]
	[ValidateSet( 'Forest', 'Sites', 'Domains', 'OUs',
		'Groups', 'GPOs', 'Misc', 'All' )]
	[String[]] $Section = 'All',

	[parameter(Mandatory=$False )] 
	[Switch]$Services=$False,
	
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)]
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(Mandatory=$False)] 
	[string]$SmtpServer="",

	[Parameter( Mandatory = $false )]
	[Switch]$SuperVerbose = $false,
	
	[parameter(Mandatory=$False)] 
	[string]$From="",

	[parameter(Mandatory=$False)] 
	[string]$To="",

	[parameter(Mandatory=$False)] 
	[Switch]$UseSSL=$False 

	)
	
#Created by Carl Webster and Michael B. Smith
#webster@carlwebster.com
#@carlwebster on Twitter
#https://www.CarlWebster.com
#
#michael@smithcons.com
#@essentialexch on Twitter
#https://www.essential.exchange/blog/

#Created on April 10, 2014

#Version 1.0 released to the community on May 31, 2014
#
#Version 2.0 is based on version 1.20
#
#Version 3.11 27-May-2022
#	Fixed bug in Function getDSUsers with MaxPasswordAge reported by Danny de Kooker
#	Moved the following section headings so that the error/warning/notice messages had a section heading
#		Domain Controllers
#		Fine-grained password policies
#
#Version 3.10 23-Apr-2022
#	Added Windows Server 2022 to AD Schema version 88
#	Fixed some text output alignment
#	In Function OutputNicItem, fixed several issues with DHCP data
#	Replaced all Get-WmiObject with Get-CimInstance
#	Some general code cleanup
#	Updated schema numbers for Exchange CUs
#		"15334" = "Exchange 2016 CU21-CU23"
#		"17003" = "Exchange 2019 CU10-CU12"
#
#Version 3.09 7-Feb-2022
#	Added to Domain Information the data for ms-DS-MachineAccountQuota
#	Changed the date format for the transcript and error log files from yyyy-MM-dd_HHmm format to the FileDateTime format
#		The format is yyyyMMddTHHmmssffff (case-sensitive, using a 4-digit year, 2-digit month, 2-digit day, 
#		the letter T as a time separator, 2-digit hour, 2-digit minute, 2-digit second, and 4-digit millisecond). 
#		For example: 20221225T0840107271.
#	Fixed the German Table of Contents (Thanks to Rene Bigler)
#		From 
#			'de-'	{ 'Automatische Tabelle 2'; Break }
#		To
#			'de-'	{ 'Automatisches Verzeichnis 2'; Break }
#	In Function AbortScript, added stopping the transcript log if the log was enabled and started
#	Replaced most script Exit calls with AbortScript to stop the transcript log if the log was enabled and started
#	Updated the help text
#	Updated the ReadMe file
#
#Version 3.08 24-Nov-2021
#	In Function AbortScript, add test for the winword process and terminate it if it is running
#	In Functions AbortScript and SaveandCloseDocumentandShutdownWord, add code from Guy Leech to test for the "Id" property before using it
#	In Function ProcessDomainControllers, added "Computer Object DN" to the output
#		If the DN doesn't contain "OU=Domain Controllers", highlight the Word/HTML output in red and add "***" to the text output
#	Updated Functions ShowScriptOptions and ProcessScriptEnd to add $ReportFooter
#	Updated schema numbers for Exchange CUs
#		"15334" = "Exchange 2016 CU21-CU22"
#		"17003" = "Exchange 2019 CU10-CU11"
#	Updated the help text
#	Updated the ReadMe file
#
#Version 3.07 11-Sep-2021
#	Added array error checking for non-empty arrays before attempting to create the Word table for most Word tables
#	Added Function OutputReportFooter
#	Added Parameter ReportFooter
#		Outputs a footer section at the end of the report.
#		Report Footer
#			Report information:
#				Created with: <Script Name> - Release Date: <Script Release Date>
#				Script version: <Script Version>
#				Started on <Date Time in Local Format>
#				Elapsed time: nn days, nn hours, nn minutes, nn.nn seconds
#				Ran from domain <Domain Name> by user <Username>
#				Ran from the folder <Folder Name>
#	Updated Function OutputADFileLocations to better report on the SYSVOL state. Code supplied by Michael B. Smith.
#	Updated Function ProcessgGPOsByOUOld to allow Word table output to handle GPOs that somehow PowerShell thinks are arrays
#	Updated Functions SaveandCloseTextDocument and SaveandCloseHTMLDocument to add a "Report Complete" line
#	Updated the help text
#	Updated the ReadMe file
#	
#Version 3.06 27-Jul-2021
#	Add new Function ProcessOUsForBlockedInheritance to add a report section for OUs with GPO Block Inheritance set
#	Add new Function ProcessSYSVOLStateInfo to show the SYSVOL state for each DC as an Appendix
#	Added by MBS, HTML codes for AlignLeft and AlignRight
#		Update Function AddHTMLTable
#		Update Function FormatHTMLTable
#		Update Function Get-ComputerCountByOS
#		Update Function getDSUsers
#		Update Function OutputEventLogInfo
#		Update Function ProcessEventLogInfo
#		Update Function ProcessGroupInformation
#		Update Function ProcessOrganizationalUnits
#		Update Function ProcessSYSVOLStateInfo
#		Update Function WriteHTMLLine
#	In Function ProcessAllDCsInTheForest, change the way all domain controllers in the forest are retrieved
#		The previous method did not always find RODC appliances
#		Use new method given by MBS
#	The following fixes were requested by Jorge de Almeida Pinto
#		In Function Get-RegistryValue, removed the Write-Verbose message on error as it confused people
#		In Function OutputADFileLocations, check only for null to catch appliances (Riverbed) with no registry
#		In Function OutputEventLogInfo, add Try/Catch to Get-EventLog to catch appliances (Riverbed) with no event logs
#		In Function OutputTimeServerRegistryKeys, check only for null to catch appliances (Riverbed) with no registry
#		When processing DCs, add testing to see if the DC is online before processing registry keys
#			Add an error message to console and output file
#		When testing a DC to see if it was online, I used the wrong variable name
#	In Function ProcessScriptEnd, always output Company Name
#	In Function ShowScriptOptions, always output Company Name
#
#Version 3.05 7-Jul-2021
#	Add fixes provided by Jorge de Almeida Pinto 
#		Fixed the way the $Script:AllDomainControllers array is built
#		Fixed getting Fine-grained Password policies to work in a multiple domain/child domain forest
#	Change the CompanyName parameter so that HTML and Text output can use it. (requested by Michael B. Smith)
#		.PARAMETER CompanyName
#			Company Name to use for the Word Cover Page or the Forest Information section for 
#			HTML and Text.
#	
#			Default value for Word output is contained in 
#			HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
#			HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
#			on the computer running the script.
#
#			This parameter has an alias of CN.
#
#			For Word output, if either registry key does not exist and this parameter is not 
#			specified, the report will not contain a Company Name on the cover page.
#	
#			For HTML and Text output, the Forest Information section will not contain the Company 
#			Name if this parameter is not specified.
#	For both HTML and Text output, at the end of the report add a "Report Complete" line (requested by Michael B. Smith)
#	For Privileged Groups, add a column for SamAccountName (requested by Michael B. Smith)
#	For the forest section, if a company name is entered, added the company name to the section title (requested by Michael B. Smith)
#	For the section Computer Operating Systems, fix the HTML tables to have slightly wider columns (requested by Michael B. Smith)
#	For Users with AdminCount=1, add columns for SamAccountName and Domain (requested by Michael B. Smith)
#	Renamed items in the list of AD Schema Items (requested by Michael B. Smith)
#		RAS Server -> NPS/RAS Server
#		LAPS -> On-premises LAPS
#		SCCM -> MECM/SCCM
#		Lync/Skype for Business -> On-premises Lync/Skype for Business
#		Exchange -> On-premises Exchange
#	Update schema numbers for Exchange CUs
#		"15333" = "Exchange 2016 CU19/CU20"
#		"15334" = "Exchange 2016 CU21"
#		"17002" = "Exchange 2019 CU8/CU9"
#		"17003" = "Exchange 2019 CU10"
#	Updated the help text
#	Updated the ReadMe file

#Version 3.04 24-Mar-2021
#	Change the wording for schema extensions from "Just because a schema extension is Present does not mean it is in use."
#		To "Just because a schema extension is Present does not mean that the product is in use."
#	Only process and output Foreign Security Principal data for the Root Domain
#	Only process the Appendix Domain Controller DNS Info if -DCDNSInfo is true. No need for an empty table and Appendix otherwise
#	Removed a few warnings from the console output that were not warnings
#	The following fixes are for running the script in a Forest with multiple domains
#		When creating the array that contains all domain controllers, don't sort after each domain as sorting changed the Type of the arraylist after the first domain was processed
#			This caused the three Appendixes to only contain the data for the DCs in the first domain
#		When outputting domain controllers, sort the DCs by domain name and DC name
#			Put the DCs in domain name order, don't put every DC in the Root domain
#			Change the header to reflect the actual domain name
#		When retrieving Inherited GPOs, add the Domain name to the cmdlet
#		When running in a child or tree domain, only the domain entered was used when calculating the number of domains in the forest
#			That is now fixed
#		When running in a child or tree domain and using -ADForest, compare the root domain's name to the name entered for -ADForest
#			If they are not the same, abort the script and state to rerun the script with -ADDomain and not -ADForest
#	Updated the help text
#	Updated the ReadMe file

#
#Version 3.03 22-Feb-2021
#	Added a Try/Catch and -LDAPFilter when checking for the Exchange schema attributes to suppress the error if Exchange is not installed
#	Added Domain SID to the Domain Information section
#	Added SYSVOL State to Function OutputADFileLocations
#		If SYSVOL State is not 4, highlight in red
#	Added updates from Michael B. Smith for MaxPasswordAge
#		Update Function getDSUsers
#		Update Function GetMaximumPasswordAge
#	Changed from using Test-Connection to Test-NetConnection -Port 88
#		Port 88 is the KDC and is unique to DCs (thanks to Matthew Woolnough for the suggestion)
#	Cleaned up console output
#	In Function BuildMultiColumnTable:
#		Prevent a division by 0 error if $MaxLength was 0
#		Fixed OutOfBounds array error (appears to be a corner case when there are 11 subnets assigned to a Site)
#	Fixed bug to now catch empty Site Subnet arrays
#		Added text "No Subnets linked to this site"
#	Updated Function GetComputerServices to add "***" in the Text output when the service type is Automatic and Status is Stopped
#	Updated Function getDSUsers to handle processing accounts in the Foreign Security Principals container
#		Find all orphaned SIDs
#		Get a count of orphaned SIDs
#		Added Function OutputFSPUserInfo to output the Orphaned SIDs and the groups those SIDs are members of
#	Updated Function ProcessGroupInformation to put HTML output in Red when:
#		Password Last Change is null or not set
#		Password Never Expires is True
#		Account is Disabled
#	Updated the help text
#	Updated the ReadMe file
#	When processing Groups for attribute adminCount -eq 1, fixed where the group name doesn’t match the samAccountName or the distinguishedName
#	When processing Groups that have attribute adminCount -eq 1, check if there was an error retrieving members of the group
#		If there was an error, add the text "Unable to retrieve group members. Check for orphaned SIDs." in place of the group members
#
#Version 3.02 9-Jan-2021
#	Added to the Computer Hardware section, the server's Power Plan
#	Changed all Write-Verbose statements from Get-Date to Get-Date -Format G as requested by Guy Leech
#	Clean up some spacing in the HTML output
#	Fixed issue with the Services and OU tables where highlighted cells were not in red
#	Fixed issue with the Word/PDF Domain Admins table with the missing Domain column
#	In Function OutputTimeServerRegistryKeys, change the call to w32tm to use Invoke-Command to get around access denied errors
#	Reordered parameters in an order recommended by Guy Leech
#	Updated Exchange schema versions to include up to Exchange 2016 CU19 and Exchange 2019 CU8
#	Updated help text
#	Updated ReadMe file
#	
#Version 3.01 12-Oct-2020
#	Handle the Domain Admins privileged group processing and output the same way the Enterprise and Schema Admins are done
#
#Version 3.00 2-Sep-2020
#	The Michael B. Smith Update and is based on version 2.22 and updated with the changes made up to 2.26
#	This is the "user/OU speedup" release. Significant efforts were spent to make the script run
#	faster in environments where large numbers of users and OUs exist.
#
#	Went to Set-StrictMode -Version Latest, from Version 2 and cleaned up all related errors
#	Rewrite AddHTMLTable, FormatHTMLTable, and WriteHTMLLine for speed and accuracy
#	Rewrite Line to use StringBuilder for speed
#	Again rewrite Line to lx for speed (not fully deployed)
#	In many places, pre-calculate the sizes of rowarray (a parameter to AddHTMLTable/FormatHTMLTable)
#		and use a fix-sized array (for speed). This caused changes in MANY places, plus several
#		foundational changes so that rowarray could be pre-calculated. This avoids creation of
#		array copies and memory thrashing. Eliminate rowarray when use is done. (More on this can
#		be done, but I believe the high-usage areas were all addressed.)
#	Replace these two incorrect Lync/SfB schema attributes
#			'msRTCSIP-UserRoutingGroupId', #Lync/SfB
#			'msRTCSIP-MirrorBackEndServer' #Lync/SfB
#		with
#			'ms-RTC-SIP-PoolAddress'
#			'ms-RTC-SIP-DomainName'
#	Stop using a Switch statement for HTML colors and use a pre-calculated HTML array (for speed).
#	Rewrite Get-RDUserSetting to GetTsAttributes (for speed)
#	Rewrite ProcessMiscDataByDomain into getDSUsers and a driver. Switch from using arraylists to List<T>.
#		Avoid array/List copies during sort. Generate a single user object shared among all lists. Stop using
#		Get-ADUser and Switch to using .NET DirectoryServices. (For large environments, memory requirements
#		have plummeted, and speed greatly increased; for small environments, the changes are likely not 
#		noticeable.) Ensure output formatting consistent among all types (Text/HTML/MSWord).
#	Update ProcessGPOSsByDomain, ProcessGPOsByOUOld, and ProcessGPOsByOUNew to only request the specific
#		info from AD that they require (still more that can be done here). Again, for speed.
#	Update OutputTimeServerRegistryKeys so that if a server isn't available, all 12 keys aren't requested.
#		That is, detect server-down on the first key request and use default values for all keys.
#	Update OutputADFileLocations for the same (don't retry if the server is known to be down)
#	Update each of the Output*UserInfo Functions so that the first parameter is Object[] instead of
#		Object. If the array contained a single element, PowerShell was unrolling it, requiring 
#		special handling. Using Object[] prevents the unrolling.
#	August 2020
#	Changed the $Section parameter to use ValidateSet().
#	In FormatHTMLTable, only write a table header line if the tableheader length is greater than zero.
#	In FormatHTMLTable, write </td> and </tr> in the proper places (previously, there weren't enough
#	</td>'s being written). I think all HTML is now "legal".
#	In FormatHTMLTable, make the docs accurate. Finally.
#	FormatHTMLTable - fix the usage of $fixedWidth and $columnIndex for good (I hope).
#	AddHTMLTable - match the usage of $fixedInfo and $columnIndex to that of FormatHTMLTable.
#	AddHTMLTable - optimize usage of $fixedInfo.
#	Further pre-calculated $rowArray rewrites.
#	Don't try too hard to analyze 'Server Core' yes or no. Was invalid check of $error array.
#	Domain Admins HTML output was missing the "Domain" column. Added.
#	All three output types were generating an error accessing TrustExtendedAttributes.TrustDirection. Fixed.
#   $DomainInfo.PublicKeyRequiredPasswordRolling could be accessed when $null. Ensure that doesn't happen.
#
#WEBSTER'S CHANGES for 3.00
#
#	Add checking for a Word version of 0, which indicates the Office installation needs repairing
#	Add Receive Side Scaling setting to Function OutputNICItem
#	Change color variables $wdColorGray15 and $wdColorGray05 from [long] to [int]
#	Change location of the -Dev, -Log, and -ScriptInfo output files from the script folder to the -Folder location (Thanks to Guy Leech for the "suggestion")
#	Change some Write-Error to Write-Warning
#	Change some Write-Warning to Write-Host
#	Change Text output to use [System.Text.StringBuilder]
#		Updated Functions Line and SaveAndCloseTextDocument
#	Fix Swedish Table of Contents (Thanks to Johan Kallio)
#		From 
#			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
#		To
#			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
#	Fixed all WriteHTMLLine lines that were supposed to be in bold. Used MBS' updates.
#	Fixed issues with the Domain Admins Privileged Group where the user type was assumed to be a User
#		Added checking for the object type and handling Groups and Users
#	Fixed issues with Word tables with later versions of PowerShell.
#	Fixed issues with Word table formatting.
#	Fixed several variable name typos
#	General code cleanup
#	HTML is now the default output format.
#	In Function OutputNicItem, change how $powerMgmt is retrieved
#		Will now show "Not Supported" instead of "N/A" if the NIC driver does not support Power Management (i.e., XenServer)
#	Reformatted the terminating Write-Error messages to make them more visible and readable in the console
#	Removed invalid URLs from the code if I could not find the original article's new location
#	Remove the SMTP parameterset and manually verify the parameters
#	Reorder parameters
#	Updated Function OutputNicItem with a $ComputerName parameter
#		Update Function GetComputerWMIInfo to pass the computer name parameter to the OutputNicItem Function
#	Update Function SetWordCellFormat to change parameter $BackgroundColor to [int]
#	Update Functions GetComputerWMIInfo and OutputNicInfo to fix two bugs in NIC Power Management settings
#	Update Function SendEmail to handle anonymous unauthenticated email
#		Update Help Text with examples
#	Updated help text
#	Updated Function SendEmail with corrections made by MBS
#	Updated the following Exchange Schema Versions:
#		"15312" = "Exchange 2013 CU7 through CU23"
#		"15317" = "Exchange 2016 Preview and RTM"
#		"15332" = "Exchange 2016 CU7 through CU15"
#		"17000" = "Exchange 2019 RTM/CU1"
#		"17001" = "Exchange 2019 CU2-CU4"
#	You can now select multiple output formats. This required extensive code changes.
#

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): System Cleanup"
		If(Test-Path variable:global:word)
		{
			$Script:Word.quit()
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
			Remove-Variable -Name word -Scope Global 4>$Null
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()

	If($MSWord -or $PDF)
	{
		#is the winword Process still running? kill it

		#find out our session (usually "1" except on TS/RDC or Citrix)
		$SessionID = (Get-Process -PID $PID).SessionId

		#Find out if winword running in our session
		$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}) | Select-Object -Property Id 
		If( $wordprocess -and $wordprocess.Id -gt 0)
		{
			Write-Verbose "$(Get-Date -Format G): WinWord Process is still running. Attempting to stop WinWord Process # $($wordprocess.Id)"
			Stop-Process $wordprocess.Id -EA 0
		}
	}
	
	Write-Verbose "$(Get-Date -Format G): Script has been aborted"
	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $True) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date -Format G): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date -Format G): Transcript/log stop failed"
			}
		}
	}
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Set-StrictMode -Version Latest

#force on
$PSDefaultParameterValues   = @{"*:Verbose"=$True}
$SaveEAPreference           = $ErrorActionPreference
$ErrorActionPreference      = 'SilentlyContinue'
$global:emailCredentials    = $Null

## v3.00
$script:ExtraSpecialVerbose = $false

#Report footer stuff
$script:MyVersion           = '3.11'
$Script:ScriptName          = "ADDS_Inventory_V3.ps1"
$tmpdate                    = [datetime] "05/27/2022"
$Script:ReleaseDate         = $tmpdate.ToUniversalTime().ToShortDateString()

Function wv
{
	$s = $args -join ''
	Write-Verbose $s
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$HTML = $True
}

If($MSWord)
{
	Write-Verbose "$(Get-Date -Format G): MSWord is set"
}
If($PDF)
{
	Write-Verbose "$(Get-Date -Format G): PDF is set"
}
If($Text)
{
	Write-Verbose "$(Get-Date -Format G): Text is set"
}
If($HTML)
{
	Write-Verbose "$(Get-Date -Format G): HTML is set"
}

If($ADForest -ne "" -and $ADDomain -ne "")
{
	#Make ADForest equal to ADDomain so no code has to change in the script
	$ADForest = $ADDomain
}

#If the MaxDetails parameter is used, set a bunch of stuff true
If($MaxDetails)
{
	$DCDNSInfo       	= $True
	$GPOInheritance  	= $True
	$HardWare        	= $True
	$IncludeUserInfo	= $True
	$Services        	= $True
	$Section			= "All"
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date -Format G): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date -Format G): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
#Do not indent the following write-error lines. Doing so will mess up the console formatting of the error message.
			Write-Error "
			`n`n
	Folder $Folder is a file, not a folder.
			`n`n
	Script cannot continue.
			`n`n"
			AbortScript
		}
	}
	Else
	{
		#does not exist
		Write-Error "
		`n`n
	Folder $Folder does not exist.
		`n`n
	Script cannot continue.
		`n`n
		"
		AbortScript
	}
}

If($Folder -eq "")
{
	$Script:pwdpath = $pwd.Path
}
Else
{
	$Script:pwdpath = $Folder
}

If($Script:pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
}

If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$Script:pwdpath\ADDSDocScriptTranscript_$(Get-Date -f FileDateTime).txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date -Format G): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Verbose "$(Get-Date -Format G): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$Script:pwdpath\ADInventoryScriptErrors_$(Get-Date -f FileDateTime).txt"
}

If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer but did not include a From or To email address.
	`n`n
	`t`t
	Script cannot Continue.
	`n`n"
	AbortScript
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`t`t
	Script cannot Continue.
	`n`n"
	AbortScript
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`t`t
	Script cannot Continue.
	`n`n"
	AbortScript
}
If(![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot Continue.
	`n`n"
	AbortScript
}
If(![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a From email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot Continue.
	`n`n"
	AbortScript
}
If(![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a To email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot Continue.
	`n`n"
	AbortScript
}

$Script:DCDNSIPInfo    = New-Object System.Collections.ArrayList
$Script:DCEventLogInfo = New-Object System.Collections.ArrayList
$Script:TimeServerInfo = New-Object System.Collections.ArrayList
$Script:DCSYSVOLState  = New-Object System.Collections.ArrayList
$Script:DARights       = $False
$Script:Elevated       = $False

#region initialize variables for word html and text
[string]$Script:RunningOS = (Get-CIMInstance -ClassName Win32_OperatingSystem -EA 0 -Verbose:$False).Caption
$Script:CoName = $CompanyName #3.05 move so this is available for HTML and Text output also

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	Write-Verbose "$(Get-Date -Format G): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorBlack = 0
	#[int]$wdColorGray05 = 15987699 
	[int]$wdColorGray15 = 14277081
	[int]$wdColorRed = 255
	[int]$wdColorWhite = 16777215
	#[int]$wdColorYellow = 65535 #added in ADDS script V2.22
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	#[int]$wdAlignParagraphLeft = 0
	#[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	#[int]$wdCellAlignVerticalCenter = 1
	#[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	#[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	#[int]$wdAdjustFirstColumn = 2
	#[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	#[int]$Indent1TabStops = 1 * $PointsPerTabStop
	#[int]$Indent2TabStops = 2 * $PointsPerTabStop
	#[int]$Indent3TabStops = 3 * $PointsPerTabStop
	#[int]$Indent4TabStops = 4 * $PointsPerTabStop

	#http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	#[int]$wdHeadingFormatFalse = 0 
}

If($HTML)
{
	#V3.00 Prior versions used Set-Variable. That hid the variables
	#from @code. So MBS Switched to using $global:

    $global:htmlredmask       = "#FF0000" 4>$Null
    $global:htmlcyanmask      = "#00FFFF" 4>$Null
    $global:htmlbluemask      = "#0000FF" 4>$Null
    $global:htmldarkbluemask  = "#0000A0" 4>$Null
    $global:htmllightbluemask = "#ADD8E6" 4>$Null
    $global:htmlpurplemask    = "#800080" 4>$Null
    $global:htmlyellowmask    = "#FFFF00" 4>$Null
    $global:htmllimemask      = "#00FF00" 4>$Null
    $global:htmlmagentamask   = "#FF00FF" 4>$Null
    $global:htmlwhitemask     = "#FFFFFF" 4>$Null
    $global:htmlsilvermask    = "#C0C0C0" 4>$Null
    $global:htmlgraymask      = "#808080" 4>$Null
    $global:htmlblackmask     = "#000000" 4>$Null
    $global:htmlorangemask    = "#FFA500" 4>$Null
    $global:htmlmaroonmask    = "#800000" 4>$Null
    $global:htmlgreenmask     = "#008000" 4>$Null
    $global:htmlolivemask     = "#808000" 4>$Null

    $global:htmlbold        = 1 4>$Null
    $global:htmlitalics     = 2 4>$Null
	$global:htmlAlignLeft   = 4 4>$Null	#3.06 added by MBS
	$global:htmlAlignRight  = 8 4>$Null	#3.06 added by MBS
    $global:htmlred         = 16 4>$Null
    $global:htmlcyan        = 32 4>$Null
    $global:htmlblue        = 64 4>$Null
    $global:htmldarkblue    = 128 4>$Null
    $global:htmllightblue   = 256 4>$Null
    $global:htmlpurple      = 512 4>$Null
    $global:htmlyellow      = 1024 4>$Null
    $global:htmllime        = 2048 4>$Null
    $global:htmlmagenta     = 4096 4>$Null
    $global:htmlwhite       = 8192 4>$Null
    $global:htmlsilver      = 16384 4>$Null
    $global:htmlgray        = 32768 4>$Null
    $global:htmlolive       = 65536 4>$Null
    $global:htmlorange      = 131072 4>$Null
    $global:htmlmaroon      = 262144 4>$Null
    $global:htmlgreen       = 524288 4>$Null
	$global:htmlblack       = 1048576 4>$Null

	$global:htmlsb          = ( $htmlsilver -bor $htmlBold ) ## point optimization

	$global:htmlColor = 
	@{
		$htmlred       = $htmlredmask
		$htmlcyan      = $htmlcyanmask
		$htmlblue      = $htmlbluemask
		$htmldarkblue  = $htmldarkbluemask
		$htmllightblue = $htmllightbluemask
		$htmlpurple    = $htmlpurplemask
		$htmlyellow    = $htmlyellowmask
		$htmllime      = $htmllimemask
		$htmlmagenta   = $htmlmagentamask
		$htmlwhite     = $htmlwhitemask
		$htmlsilver    = $htmlsilvermask
		$htmlgray      = $htmlgraymask
		$htmlolive     = $htmlolivemask
		$htmlorange    = $htmlorangemask
		$htmlmaroon    = $htmlmaroonmask
		$htmlgreen     = $htmlgreenmask
		$htmlblack     = $htmlblackmask
	}
}
#endregion

#region email Function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Verbose "$(Get-Date -Format G): Prepare to email"

	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.

"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()
	
	If($From -Like "anonymous@*")
	{
		#https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
		$anonUsername = "anonymous"
		$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
		$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $anonCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $anonCredentials *>$Null 
		}
		
		If($?)
		{
			Write-Verbose "$(Get-Date -Format G): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date -Format G): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Verbose "$(Get-Date -Format G): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
		}

		If(!$?)
		{
			$e = $error[0]
			
			#error 5.7.57 is O365 and error 5.7.0 is gmail
			If($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7"))
			{
				#The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
				Write-Verbose "$(Get-Date -Format G): Current user's credentials failed. Ask for usable credentials."

				If($Dev)
				{
					Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
				}

				$error.Clear()

				$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

				If($UseSSL)
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-UseSSL -credential $emailCredentials *>$Null 
				}
				Else
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-credential $emailCredentials *>$Null 
				}

				If($?)
				{
					Write-Verbose "$(Get-Date -Format G): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date -Format G): Email was not sent:"
					Write-Warning "$(Get-Date): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): Email was not sent:"
				Write-Warning "$(Get-Date): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	
	# modified 17-Aug-2016 to fix a few issues with Text and HTML output
	# modified 29-Apr-2018 to change from Arrays to New-Object System.Collections.ArrayList
	# modified 11-Mar-2022 changed from using Get-WmiObject to Get-CimInstance

	#Get Computer info
	Write-Verbose "$(Get-Date -Format G): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date -Format G): `t`t`tHardware information"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteWordLine 4 0 "General Computer"
	}
	If($Text)
	{
		Line 0 "Computer Information: $($RemoteComputerName)"
		Line 1 "General Computer"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteHTMLLine 4 0 "General Computer"
	}
	
	Try
	{
		If($RemoteComputerName -eq $env:computername)
		{
			$Results = Get-CimInstance -ClassName win32_computersystem -Verbose:$False
		}
		Else
		{
			$Results = Get-CimInstance -computername $RemoteComputerName -ClassName win32_computersystem -Verbose:$False
		}
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Null -ne $Results)
	{
		$ComputerItems = $Results | Select-Object Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
		$Results = $Null
		If($RemoteComputerName -eq $env:computername)
		{
			[string]$ComputerOS = (Get-CimInstance -ClassName Win32_OperatingSystem -EA 0 -Verbose:$False).Caption
		}
		Else
		{
			[string]$ComputerOS = (Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $RemoteComputerName -EA 0 -Verbose:$False).Caption
		}

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item $ComputerOS $RemoteComputerName
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date -Format G): Get-CimInstance win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-CimInstance win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-CimInstance win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 3 "Get-CimInstance win32_computersystem failed for $($RemoteComputerName)"
			Line 3 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-CimInstance win32_computersystem failed for $($RemoteComputerName)" -option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 3 "No results Returned for Computer information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Computer information" -Option $htmlBold
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date -Format G): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Drive(s)"
	}
	If($Text)
	{
		Line 1 "Drive(s)"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Drive(s)"
	}

	Try
	{
		If($RemoteComputerName -eq $env:computername)
		{
			$Results = Get-CimInstance -ClassName Win32_LogicalDisk -Verbose:$False
		}
		Else
		{
			$Results = Get-CimInstance -CimSession $RemoteComputerName -ClassName Win32_LogicalDisk -Verbose:$False
		}
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$drives = $Results | Select-Object caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date -Format G): Get-CimInstance Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-CimInstance Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-CimInstance Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Get-CimInstance Win32_LogicalDisk failed for $($RemoteComputerName)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-CimInstance Win32_LogicalDisk failed for $($RemoteComputerName)" -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for Drive information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Drive information" -Option $htmlBold
		}
	}
	
	#Get CPU's and stepping
	Write-Verbose "$(Get-Date -Format G): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Processor(s)"
	}
	If($Text)
	{
		Line 1 "Processor(s)"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Processor(s)"
	}

	Try
	{
		If($RemoteComputerName -eq $env:computername)
		{
			$Results = Get-CimInstance -ClassName win32_Processor -Verbose:$False
		}
		Else
		{
			$Results = Get-CimInstance -computername $RemoteComputerName -ClassName win32_Processor -Verbose:$False
		}
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Processors = $Results | Select-Object availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date -Format G): Get-CimInstance win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-CimInstance win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-CimInstance win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Get-CimInstance win32_Processor failed for $($RemoteComputerName)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-CimInstance win32_Processor failed for $($RemoteComputerName)" -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for Processor information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Processor information" -Option $htmlBold
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date -Format G): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Interface(s)"
	}
	If($Text)
	{
		Line 1 "Network Interface(s)"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Network Interface(s)"
	}

	[bool]$GotNics = $True
	
	Try
	{
		If($RemoteComputerName -eq $env:computername)
		{
			$Results = Get-CimInstance -ClassName win32_networkadapterconfiguration -Verbose:$False
		}
		Else
		{
			$Results = Get-CimInstance -computername $RemoteComputerName -ClassName win32_networkadapterconfiguration -Verbose:$False
		}
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Nics = $Results | Where-Object {$Null -ne $_.ipaddress}
		$Results = $Null

		If($Null -eq $Nics) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = $True
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					If($RemoteComputerName -eq $env:computername)
					{
						$ThisNic = Get-CimInstance -ClassName win32_networkadapter -Verbose:$False | Where-Object {$_.index -eq $nic.index}
					}
					Else
					{
						$ThisNic = Get-CimInstance -computername $RemoteComputerName -ClassName win32_networkadapter -Verbose:$False | Where-Object {$_.index -eq $nic.index}
					}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $Null -ne $ThisNic)
				{
					OutputNicItem $Nic $ThisNic $RemoteComputerName
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose "$(Get-Date -Format G): Get-CimInstance win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-CimInstance win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
					}
					If($Text)
					{
						Line 2 "Error retrieving NIC information"
					}
					If($HTML)
					{
						WriteHTMLLine 0 2 "Error retrieving NIC information" -Option $htmlBold
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date -Format G): No results Returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
					If($Text)
					{
						Line 2 "No results Returned for NIC information"
					}
					If($HTML)
					{
						WriteHTMLLine 0 2 "No results Returned for NIC information" -Option $htmlBold
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date -Format G): Get-CimInstance win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-CimInstance win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Error retrieving NIC configuration information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Error retrieving NIC configuration information" -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" -Option $htmlBold
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 0 ""
	}
}

Function OutputComputerItem
{
	Param([object]$Item, [string]$OS, [string]$RemoteComputerName)
	
	#get computer's power plan
	#https://techcommunity.microsoft.com/t5/core-infrastructure-and-security/get-the-active-power-plan-of-multiple-servers-with-powershell/ba-p/370429
	
	try 
	{

		If($RemoteComputerName -eq $env:computername)
		{
			$PowerPlan = (Get-CimInstance -ClassName Win32_PowerPlan -Namespace "root\cimv2\power" -Verbose:$False |
				Where-Object {$_.IsActive -eq $true} |
				Select-Object @{Name = "PowerPlan"; Expression = {$_.ElementName}}).PowerPlan
		}
		Else
		{
			$PowerPlan = (Get-CimInstance -CimSession $RemoteComputerName -ClassName Win32_PowerPlan -Namespace "root\cimv2\power" -Verbose:$False |
				Where-Object {$_.IsActive -eq $true} |
				Select-Object @{Name = "PowerPlan"; Expression = {$_.ElementName}}).PowerPlan
		}
	}

	catch 
	{

		$PowerPlan = $_.Exception

	}	
	
	If($MSWord -or $PDF)
	{
		$ItemInformation = New-Object System.Collections.ArrayList
		$ItemInformation.Add(@{ Data = "Manufacturer"; Value = $Item.manufacturer; }) > $Null
		$ItemInformation.Add(@{ Data = "Model"; Value = $Item.model; }) > $Null
		$ItemInformation.Add(@{ Data = "Domain"; Value = $Item.domain; }) > $Null
		$ItemInformation.Add(@{ Data = "Operating System"; Value = $OS; }) > $Null
		$ItemInformation.Add(@{ Data = "Power Plan"; Value = $PowerPlan; }) > $Null
		$ItemInformation.Add(@{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }) > $Null
		$ItemInformation.Add(@{ Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }) > $Null
		$ItemInformation.Add(@{ Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }) > $Null
		$Table = AddWordTable -Hashtable $ItemInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Manufacturer`t`t`t: " $Item.manufacturer
		Line 2 "Model`t`t`t`t: " $Item.model
		Line 2 "Domain`t`t`t`t: " $Item.domain
		Line 2 "Operating System`t`t: " $OS
		Line 2 "Power Plan`t`t`t: " $PowerPlan
		Line 2 "Total Ram`t`t`t: $($Item.totalphysicalram) GB"
		Line 2 "Physical Processors (sockets)`t: " $Item.NumberOfProcessors
		Line 2 "Logical Processors (cores w/HT)`t: " $Item.NumberOfLogicalProcessors
		Line 2 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Manufacturer",($global:htmlsb),$Item.manufacturer,$htmlwhite)
		$rowdata += @(,('Model',($global:htmlsb),$Item.model,$htmlwhite))
		$rowdata += @(,('Domain',($global:htmlsb),$Item.domain,$htmlwhite))
		$rowdata += @(,('Operating System',($global:htmlsb),$OS,$htmlwhite))
		$rowdata += @(,('Power Plan',($global:htmlsb),$PowerPlan,$htmlwhite))
		$rowdata += @(,('Total Ram',($global:htmlsb),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($global:htmlsb),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($global:htmlsb),$Item.NumberOfLogicalProcessors,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	
	$xDriveType = ""
	Switch ($drive.drivetype)
	{
		0		{$xDriveType = "Unknown"; Break}
		1		{$xDriveType = "No Root Directory"; Break}
		2		{$xDriveType = "Removable Disk"; Break}
		3		{$xDriveType = "Local Disk"; Break}
		4		{$xDriveType = "Network Drive"; Break}
		5		{$xDriveType = "Compact Disc"; Break}
		6		{$xDriveType = "RAM Disk"; Break}
		Default {$xDriveType = "Unknown"; Break}
	}
	
	$xVolumeDirty = ""
	If(![String]::IsNullOrEmpty($drive.volumedirty))
	{
		If($drive.volumedirty)
		{
			$xVolumeDirty = "Yes"
		}
		Else
		{
			$xVolumeDirty = "No"
		}
	}

	If($MSWORD -or $PDF)
	{
		$DriveInformation = New-Object System.Collections.ArrayList
		$DriveInformation.Add(@{ Data = "Caption"; Value = $Drive.caption; }) > $Null
		$DriveInformation.Add(@{ Data = "Size"; Value = "$($drive.drivesize) GB"; }) > $Null
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation.Add(@{ Data = "File System"; Value = $Drive.filesystem; }) > $Null
		}
		$DriveInformation.Add(@{ Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }) > $Null
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation.Add(@{ Data = "Volume Name"; Value = $Drive.volumename; }) > $Null
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$DriveInformation.Add(@{ Data = "Volume is Dirty"; Value = $xVolumeDirty; }) > $Null
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation.Add(@{ Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }) > $Null
		}
		$DriveInformation.Add(@{ Data = "Drive Type"; Value = $xDriveType; }) > $Null
		$Table = AddWordTable -Hashtable $DriveInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells `
		-Bold `
		-BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	If($Text)
	{
		Line 2 "Caption`t`t: " $drive.caption
		Line 2 "Size`t`t: $($drive.drivesize) GB"
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			Line 2 "File System`t: " $drive.filesystem
		}
		Line 2 "Free Space`t: $($drive.drivefreespace) GB"
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			Line 2 "Volume Name`t: " $drive.volumename
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			Line 2 "Volume is Dirty`t: " $xVolumeDirty
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			Line 2 "Volume Serial #`t: " $drive.volumeserialnumber
		}
		Line 2 "Drive Type`t: " $xDriveType
		Line 2 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Caption",($global:htmlsb),$Drive.caption,$htmlwhite)
		$rowdata += @(,('Size',($global:htmlsb),"$($drive.drivesize) GB",$htmlwhite))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',($global:htmlsb),$Drive.filesystem,$htmlwhite))
		}
		$rowdata += @(,('Free Space',($global:htmlsb),"$($drive.drivefreespace) GB",$htmlwhite))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',($global:htmlsb),$Drive.volumename,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',($global:htmlsb),$xVolumeDirty,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',($global:htmlsb),$Drive.volumeserialnumber,$htmlwhite))
		}
		$rowdata += @(,('Drive Type',($global:htmlsb),$xDriveType,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1		{$xAvailability = "Other"; Break}
		2		{$xAvailability = "Unknown"; Break}
		3		{$xAvailability = "Running or Full Power"; Break}
		4		{$xAvailability = "Warning"; Break}
		5		{$xAvailability = "In Test"; Break}
		6		{$xAvailability = "Not Applicable"; Break}
		7		{$xAvailability = "Power Off"; Break}
		8		{$xAvailability = "Off Line"; Break}
		9		{$xAvailability = "Off Duty"; Break}
		10		{$xAvailability = "Degraded"; Break}
		11		{$xAvailability = "Not Installed"; Break}
		12		{$xAvailability = "Install Error"; Break}
		13		{$xAvailability = "Power Save - Unknown"; Break}
		14		{$xAvailability = "Power Save - Low Power Mode"; Break}
		15		{$xAvailability = "Power Save - Standby"; Break}
		16		{$xAvailability = "Power Cycle"; Break}
		17		{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	If($MSWORD -or $PDF)
	{
		$ProcessorInformation = New-Object System.Collections.ArrayList
		$ProcessorInformation.Add(@{ Data = "Name"; Value = $Processor.name; }) > $Null
		$ProcessorInformation.Add(@{ Data = "Description"; Value = $Processor.description; }) > $Null
		$ProcessorInformation.Add(@{ Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }) > $Null
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }) > $Null
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }) > $Null
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "Number of Cores"; Value = $Processor.numberofcores; }) > $Null
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "Number of Logical Processors (cores w/HT)"; Value = $Processor.numberoflogicalprocessors; }) > $Null
		}
		$ProcessorInformation.Add(@{ Data = "Availability"; Value = $xAvailability; }) > $Null
		$Table = AddWordTable -Hashtable $ProcessorInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Name`t`t`t`t: " $processor.name
		Line 2 "Description`t`t`t: " $processor.description
		Line 2 "Max Clock Speed`t`t`t: $($processor.maxclockspeed) MHz"
		If($processor.l2cachesize -gt 0)
		{
			Line 2 "L2 Cache Size`t`t`t: $($processor.l2cachesize) KB"
		}
		If($processor.l3cachesize -gt 0)
		{
			Line 2 "L3 Cache Size`t`t`t: $($processor.l3cachesize) KB"
		}
		If($processor.numberofcores -gt 0)
		{
			Line 2 "# of Cores`t`t`t: " $processor.numberofcores
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			Line 2 "# of Logical Procs (cores w/HT)`t: " $processor.numberoflogicalprocessors
		}
		Line 2 "Availability`t`t`t: " $xAvailability
		Line 2 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($global:htmlsb),$Processor.name,$htmlwhite)
		$rowdata += @(,('Description',($global:htmlsb),$Processor.description,$htmlwhite))

		$rowdata += @(,('Max Clock Speed',($global:htmlsb),"$($processor.maxclockspeed) MHz",$htmlwhite))
		If($processor.l2cachesize -gt 0)
		{
			$rowdata += @(,('L2 Cache Size',($global:htmlsb),"$($processor.l2cachesize) KB",$htmlwhite))
		}
		If($processor.l3cachesize -gt 0)
		{
			$rowdata += @(,('L3 Cache Size',($global:htmlsb),"$($processor.l3cachesize) KB",$htmlwhite))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',($global:htmlsb),$Processor.numberofcores,$htmlwhite))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',($global:htmlsb),$Processor.numberoflogicalprocessors,$htmlwhite))
		}
		$rowdata += @(,('Availability',($global:htmlsb),$xAvailability,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic, [string]$RemoteComputerName)
	
	If($RemoteComputerName -eq $env:computername)
	{
		$powerMgmt = Get-CimInstance -ClassName MSPower_DeviceEnable -Namespace "root\wmi" -Verbose:$False |
			Where-Object{$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}
	}
	Else
	{
		$powerMgmt = Get-CimInstance -CimSession $RemoteComputerName -ClassName MSPower_DeviceEnable -Namespace "root\wmi" -Verbose:$False |
			Where-Object{$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}
	}

	If($? -and $Null -ne $powerMgmt)
	{
		If($powerMgmt.Enable -eq $True)
		{
			$PowerSaving = "Enabled"
		}
		Else	
		{
			$PowerSaving = "Disabled"
		}
	}
	Else
	{
        $PowerSaving = "N/A"
	}
	
	$xAvailability = ""
	Switch ($ThisNic.availability)
	{
		1		{$xAvailability = "Other"; Break}
		2		{$xAvailability = "Unknown"; Break}
		3		{$xAvailability = "Running or Full Power"; Break}
		4		{$xAvailability = "Warning"; Break}
		5		{$xAvailability = "In Test"; Break}
		6		{$xAvailability = "Not Applicable"; Break}
		7		{$xAvailability = "Power Off"; Break}
		8		{$xAvailability = "Off Line"; Break}
		9		{$xAvailability = "Off Duty"; Break}
		10		{$xAvailability = "Degraded"; Break}
		11		{$xAvailability = "Not Installed"; Break}
		12		{$xAvailability = "Install Error"; Break}
		13		{$xAvailability = "Power Save - Unknown"; Break}
		14		{$xAvailability = "Power Save - Low Power Mode"; Break}
		15		{$xAvailability = "Power Save - Standby"; Break}
		16		{$xAvailability = "Power Cycle"; Break}
		17		{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	#attempt to get Receive Side Scaling setting
	$RSSEnabled = "N/A"
	Try
	{
		#https://ios.developreference.com/article/10085450/How+do+I+enable+VRSS+(Virtual+Receive+Side+Scaling)+for+a+Windows+VM+without+relying+on+Enable-NetAdapterRSS%3F
		If($RemoteComputerName -eq $env:computername)
		{
			$RSSEnabled = (Get-CimInstance -ClassName MSFT_NetAdapterRssSettingData -Namespace "root\StandardCimV2" -ea 0 -Verbose:$False).Enabled
		}
		Else
		{
			$RSSEnabled = (Get-CimInstance -CimSession $RemoteComputerName -ClassName MSFT_NetAdapterRssSettingData -Namespace "root\StandardCimV2" -ea 0 -Verbose:$False).Enabled
		}

		If($RSSEnabled)
		{
			$RSSEnabled = "Enabled"
		}
		Else
		{
			$RSSEnabled = "Disabled"
		}
	}
	
	Catch
	{
		$RSSEnabled = "Unable to determine for $RemoteComputerName"
	}

	$xIPAddress = New-Object System.Collections.ArrayList
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress.Add("$($IPAddress)") > $Null
	}

	$xIPSubnet = New-Object System.Collections.ArrayList
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet.Add("$($IPSubnet)") > $Null
	}

	If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = New-Object System.Collections.ArrayList
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder.Add("$($DNSDomain)") > $Null
		}
	}
	
	If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = New-Object System.Collections.ArrayList
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder.Add("$($DNSServer)") > $Null
		}
	}

	$xdnsenabledforwinsresolution = ""
	If($nic.dnsenabledforwinsresolution)
	{
		$xdnsenabledforwinsresolution = "Yes"
	}
	Else
	{
		$xdnsenabledforwinsresolution = "No"
	}
	
	$xTcpipNetbiosOptions = ""
	Switch ($nic.TcpipNetbiosOptions)
	{
		0		{$xTcpipNetbiosOptions = "Use NetBIOS setting from DHCP Server"; Break}
		1		{$xTcpipNetbiosOptions = "Enable NetBIOS"; Break}
		2		{$xTcpipNetbiosOptions = "Disable NetBIOS"; Break}
		Default	{$xTcpipNetbiosOptions = "Unknown"; Break}
	}
	
	$xwinsenablelmhostslookup = ""
	If($nic.winsenablelmhostslookup)
	{
		$xwinsenablelmhostslookup = "Yes"
	}
	Else
	{
		$xwinsenablelmhostslookup = "No"
	}

	If($nic.dhcpenabled)
	{
		$DHCPLeaseObtainedDate = $nic.dhcpleaseobtained.ToLocalTime()
		If($nic.DHCPLeaseExpires -lt $nic.DHCPLeaseObtained)
		{
			#Could be an Azure DHCP Lease
			$DHCPLeaseExpiresDate = (Get-Date).AddSeconds([UInt32]::MaxValue).ToLocalTime()
		}
		Else
		{
			$DHCPLeaseExpiresDate = $nic.DHCPLeaseExpires.ToLocalTime()
		}
	}
		
	If($MSWORD -or $PDF)
	{
		$NicInformation = New-Object System.Collections.ArrayList
		$NicInformation.Add(@{ Data = "Name"; Value = $ThisNic.Name; }) > $Null
		If($ThisNic.Name -ne $nic.description)
		{
			$NicInformation.Add(@{ Data = "Description"; Value = $Nic.description; }) > $Null
		}
		$NicInformation.Add(@{ Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }) > $Null
		If(validObject $Nic Manufacturer)
		{
			$NicInformation.Add(@{ Data = "Manufacturer"; Value = $Nic.manufacturer; }) > $Null
		}
		$NicInformation.Add(@{ Data = "Availability"; Value = $xAvailability; }) > $Null
		$NicInformation.Add(@{ Data = "Allow the computer to turn off this device to save power"; Value = $PowerSaving; }) > $Null
		$NicInformation.Add(@{ Data = "Receive Side Scaling"; Value = $RSSEnabled; }) > $Null
		$NicInformation.Add(@{ Data = "Physical Address"; Value = $Nic.macaddress; }) > $Null
		If($xIPAddress.Count -gt 1)
		{
			$NicInformation.Add(@{ Data = "IP Address"; Value = $xIPAddress[0]; }) > $Null
			$NicInformation.Add(@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) > $Null
			$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = "IP Address"; Value = $tmp; }) > $Null
					$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet[$cnt]; }) > $Null
				}
			}
		}
		Else
		{
			$NicInformation.Add(@{ Data = "IP Address"; Value = $xIPAddress; }) > $Null
			$NicInformation.Add(@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) > $Null
			$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet; }) > $Null
		}
		If($nic.dhcpenabled)
		{
			$NicInformation.Add(@{ Data = "DHCP Enabled"; Value = $Nic.dhcpenabled.ToString(); }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Server"; Value = $Nic.dhcpserver; }) > $Null
		}
		Else
		{
			$NicInformation.Add(@{ Data = "DHCP Enabled"; Value = $Nic.dhcpenabled.ToString(); }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation.Add(@{ Data = "DNS Domain"; Value = $Nic.dnsdomain; }) > $Null
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$NicInformation.Add(@{ Data = "DNS Search Suffixes"; Value = $xnicdnsdomainsuffixsearchorder[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = ""; Value = $tmp; }) > $Null
				}
			}
		}
		$NicInformation.Add(@{ Data = "DNS WINS Enabled"; Value = $xdnsenabledforwinsresolution; }) > $Null
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$NicInformation.Add(@{ Data = "DNS Servers"; Value = $xnicdnsserversearchorder[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = ""; Value = $tmp; }) > $Null
				}
			}
		}
		$NicInformation.Add(@{ Data = "NetBIOS Setting"; Value = $xTcpipNetbiosOptions; }) > $Null
		$NicInformation.Add(@{ Data = "WINS: Enabled LMHosts"; Value = $xwinsenablelmhostslookup; }) > $Null
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation.Add(@{ Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation.Add(@{ Data = "Primary Server"; Value = $Nic.winsprimaryserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation.Add(@{ Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation.Add(@{ Data = "Scope ID"; Value = $Nic.winsscopeid; }) > $Null
		}
		$Table = AddWordTable -Hashtable $NicInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Name`t`t`t: " $ThisNic.Name
		If($ThisNic.Name -ne $nic.description)
		{
			Line 2 "Description`t`t: " $nic.description
		}
		Line 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
		If(validObject $Nic Manufacturer)
		{
			Line 2 "Manufacturer`t`t: " $Nic.manufacturer
		}
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 "Allow computer to turn "
		Line 2 "off device to save power: " $PowerSaving
		Line 2 "Physical Address`t: " $nic.macaddress
		Line 2 "Receive Side Scaling`t: " $RSSEnabled
		Line 2 "IP Address`t`t: " $xIPAddress[0]
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}
		Line 2 "Default Gateway`t`t: " $Nic.Defaultipgateway
		Line 2 "Subnet Mask`t`t: " $xIPSubnet[0]
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}
		If($nic.dhcpenabled)
		{
			Line 2 "DHCP Enabled`t`t: " $nic.dhcpenabled.ToString()
			Line 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
			Line 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
			Line 2 "DHCP Server`t`t: " $nic.dhcpserver
		}
		Else
		{
			Line 2 "DHCP Enabled`t`t: " $nic.dhcpenabled.ToString()
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Line 2 "DNS Domain`t`t: " $nic.dnsdomain
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Search Suffixes`t: " $xnicdnsdomainsuffixsearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
		}
		Line 2 "DNS WINS Enabled`t: " $xdnsenabledforwinsresolution
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Servers`t`t: " $xnicdnsserversearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
		}
		Line 2 "NetBIOS Setting`t`t: " $xTcpipNetbiosOptions
		Line 2 "WINS:"
		Line 3 "Enabled LMHosts`t: " $xwinsenablelmhostslookup
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			Line 3 "Host Lookup File`t: " $nic.winshostlookupfile
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			Line 3 "Primary Server`t: " $nic.winsprimaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			Line 3 "Secondary Server`t: " $nic.winssecondaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			Line 3 "Scope ID`t`t: " $nic.winsscopeid
		}
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($global:htmlsb),$ThisNic.Name,$htmlwhite)
		If($ThisNic.Name -ne $nic.description)
		{
			$rowdata += @(,('Description',($global:htmlsb),$Nic.description,$htmlwhite))
		}
		$rowdata += @(,('Connection ID',($global:htmlsb),$ThisNic.NetConnectionID,$htmlwhite))
		If(validObject $Nic Manufacturer)
		{
			$rowdata += @(,('Manufacturer',($global:htmlsb),$Nic.manufacturer,$htmlwhite))
		}
		$rowdata += @(,('Availability',($global:htmlsb),$xAvailability,$htmlwhite))
		$rowdata += @(,('Allow the computer to turn off this device to save power',($global:htmlsb),$PowerSaving,$htmlwhite))
		$rowdata += @(,('Physical Address',($global:htmlsb),$Nic.macaddress,$htmlwhite))
		$rowdata += @(,('Receive Side Scaling',($global:htmlsb),$RSSEnabled,$htmlwhite))
		$rowdata += @(,('IP Address',($global:htmlsb),$xIPAddress[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('IP Address',($global:htmlsb),$tmp,$htmlwhite))
			}
		}
		$rowdata += @(,('Default Gateway',($global:htmlsb),$Nic.Defaultipgateway[0],$htmlwhite))
		$rowdata += @(,('Subnet Mask',($global:htmlsb),$xIPSubnet[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('Subnet Mask',($global:htmlsb),$tmp,$htmlwhite))
			}
		}
		If($nic.dhcpenabled)
		{
			$rowdata += @(,('DHCP Enabled',($global:htmlsb),$Nic.dhcpenabled.ToString(),$htmlwhite))
			$rowdata += @(,('DHCP Lease Obtained',($global:htmlsb),$dhcpleaseobtaineddate,$htmlwhite))
			$rowdata += @(,('DHCP Lease Expires',($global:htmlsb),$dhcpleaseexpiresdate,$htmlwhite))
			$rowdata += @(,('DHCP Server',($global:htmlsb),$Nic.dhcpserver,$htmlwhite))
		}
		Else
		{
			$rowdata += @(,('DHCP Enabled',($global:htmlsb),$Nic.dhcpenabled.ToString(),$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$rowdata += @(,('DNS Domain',($global:htmlsb),$Nic.dnsdomain,$htmlwhite))
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Search Suffixes',($global:htmlsb),$xnicdnsdomainsuffixsearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('DNS WINS Enabled',($global:htmlsb),$xdnsenabledforwinsresolution,$htmlwhite))
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',($global:htmlsb),$xnicdnsserversearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($global:htmlsb),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('NetBIOS Setting',($global:htmlsb),$xTcpipNetbiosOptions,$htmlwhite))
		$rowdata += @(,('WINS: Enabled LMHosts',($global:htmlsb),$xwinsenablelmhostslookup,$htmlwhite))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',($global:htmlsb),$Nic.winshostlookupfile,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',($global:htmlsb),$Nic.winsprimaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',($global:htmlsb),$Nic.winssecondaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',($global:htmlsb),$Nic.winsscopeid,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region GetComputerServices
Function GetComputerServices 
{
	Param([string]$RemoteComputerName)
	# modified 29-Apr-2018 to change from Arrays to New-Object System.Collections.ArrayList
	
	#Get Computer services info
	Write-Verbose "$(Get-Date -Format G): `t`tProcessing Computer services information"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 "Services"
	}
	If($Text)
	{
		Line 0 "Services"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Services"
	}

	Try
	{
		#Iain Brighton optimization 5-Jun-2014
		#Replaced with a single call to retrieve services via WMI. The repeated
		## "Get-WMIObject Win32_Service -Filter" calls were the major delays in the script.
		## If we need to retrieve the StartUp type might as well just use WMI.
		
		#$Services = @(Get-WMIObject Win32_Service -ComputerName $RemoteComputerName | Sort-Object DisplayName)
		If($RemoteComputerName -eq $env:computername)
		{
			$Services = @(Get-CIMInstance Win32_Service -EA 0 -Verbose:$False | Sort-Object DisplayName)
		}
		Else
		{
			$Services = @(Get-CIMInstance Win32_Service -CIMSession $RemoteComputerName -EA 0 -Verbose:$False | Sort-Object DisplayName)
		}
	}
	
	Catch
	{
		$Services = $Null
	}
	
	If($? -and $Null -ne $Services)
	{
		[int]$NumServices = $Services.count
		Write-Verbose "$(Get-Date -Format G): `t`t`t$($NumServices) Services found"

		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Services ($NumServices Services found)"

			$ServicesWordTable = New-Object System.Collections.ArrayList
			$WordHighlightedCells = New-Object System.Collections.ArrayList
			[int] $CurrentServiceIndex = 2
		}
		If($Text)
		{
			Line 0 "Services ($NumServices Services found)"
			Line 0 ""
			
			[int]$MaxDisplayNameLength = ($Services.DisplayName | Measure-Object -Maximum -Property Length).Maximum
			If($MaxDisplayNameLength -gt 12) #12 is length of "Display Name"
			{
				#10 is length of "Display Name" minus 2 to allow for spacing between columns
				Line 1 ("Display Name" + (' ' * ($MaxDisplayNameLength - 10))) -NoNewLine
			}
			Else
			{
				Line 1 "Display Name " -NoNewLine
			}
			Line 1 "Status  " -NoNewLine
			Line 1 "Startup Type"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "Services ($NumServices Services found)"
			#V3.00 rowdata is pre-allocated
			$rowData = New-Object System.Array[] $Services.Count
			$rowIndx = 0
		}

		ForEach($Service in $Services) 
		{
			If($MSWord -or $PDF)
			{

				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ 
				DisplayName = $Service.DisplayName; 
				Status = $Service.State; 
				StartMode = $Service.StartMode
				}

				## Add the hash to the array
				$ServicesWordTable.Add($WordTableRowHash) > $Null

				## Store "to highlight" cell references
				If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
				{
					$WordHighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 2; }) > $Null
				}
				$CurrentServiceIndex++;
			}
			If($Text)
			{
				If(($Service.DisplayName).Length -lt ($MaxDisplayNameLength))
				{
					[int]$NumOfSpaces = (($MaxDisplayNameLength) - ($Service.DisplayName.Length)) + 2 #+2 to allow for column spacing
					$tmp1 = ($($Service.DisplayName) + (' ' * $NumOfSpaces))
					Line 1 $tmp1 -NoNewLine
				}
				Else
				{
					Line 1 "$($Service.DisplayName)  " -NoNewLine
				}
				If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
				{
					Line 1 "***$($Service.State)*** " -NoNewLine
				}
				Else
				{
					Line 1 "$($Service.State) " -NoNewLine
				}
				Line 1 $Service.StartMode
				
			}
			If($HTML)
			{
				If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
				{
					$HTMLHighlightedCells = $htmlred
				}
				Else
				{
					$HTMLHighlightedCells = $htmlwhite
				} 
				$rowdata[ $rowIndx ] = @(,($Service.DisplayName,$htmlwhite,
								$Service.State,$HTMLHighlightedCells,
								$Service.StartMode,$htmlwhite))
				$rowIndx++
			}
		}

		If($MSWord -or $PDF)
		{
			If($ServicesWordTable.Count -gt 0)	#3.07
			{
				$Table = AddWordTable -Hashtable $ServicesWordTable `
				-Columns DisplayName, Status, StartMode `
				-Headers "Display Name", "Status", "Startup Type" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
				If($WordHighlightedCells.Count -gt 0)
				{
					SetWordCellFormat -Coordinates $WordHighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;
				}

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
				'Display Name',$htmlsb,
				'Status',$htmlsb,
				'Startup Type',$htmlsb
			)
			FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 ''
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "No services were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $False $True
			WriteWordLine 0 1 "If this is a trusted Forest, you may need to rerun the" "" $Null 0 $False $True
			WriteWordLine 0 1 "script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 0 "Warning: No Services were retrieved"
			Line 1 "If this is a trusted Forest, you may need to rerun the"
			Line 1 "script with Domain Admin credentials from the trusted Forest."
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Services were retrieved" -options $htmlBold
			WriteHTMLLine 0 1 "If this is a trusted Forest, you may need to rerun the" -options $htmlBold
			WriteHTMLLine 0 1 "script with Domain Admin credentials from the trusted Forest." -options $htmlBold
		}
	}
	Else
	{
		Write-Warning "Services retrieval was successful but no services were Returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Services retrieval was successful but no services were Returned." "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 0 "Services retrieval was successful but no services were Returned."
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Services retrieval was successful but no services were Returned." -options $htmlBold
		}
	}
}
#endregion

#region BuildDCDNSIPConfigTable
Function BuildDCDNSIPConfigTable
{
	Param([string]$RemoteComputerName, [string]$Site)
	
	[bool]$GotNics = $True
	
	Try
	{
		#$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
		If($RemoteComputerName -eq $env:computername)
		{
			$Results = Get-CimInstance -ClassName win32_networkadapterconfiguration -Verbose:$False
		}
		Else
		{
			$Results = Get-CimInstance -computername $RemoteComputerName -ClassName win32_networkadapterconfiguration -Verbose:$False
		}
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Nics = $Results | Where-Object {$Null -ne $_.ipaddress}
		$Results = $Null

		If($Null -eq $Nics) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = $True
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					#$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where-Object {$_.index -eq $nic.index}
					If($RemoteComputerName -eq $env:computername)
					{
						$ThisNic = Get-CimInstance -ClassName win32_networkadapter -Verbose:$False | Where-Object {$_.index -eq $nic.index}
					}
					Else
					{
						$ThisNic = Get-CimInstance -computername $RemoteComputerName -ClassName win32_networkadapter -Verbose:$False | Where-Object {$_.index -eq $nic.index}
					}
				}
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $Null -ne $ThisNic)
				{
					Write-Verbose "$(Get-Date -Format G): `t`t`tGather DC DNS IP Config info"
					$xIPAddress = @()
					ForEach($IPAddress in $Nic.ipaddress)
					{
						$xIPAddress += "$($IPAddress)"
					}
					
					If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
					{
						$nicdnsserversearchorder = $nic.dnsserversearchorder
						$xnicdnsserversearchorder = @()
						ForEach($DNSServer in $nicdnsserversearchorder)
						{
							$xnicdnsserversearchorder += "$($DNSServer)"
						}
					}

					$obj = New-Object -TypeName PSObject
					$obj | Add-Member -MemberType NoteProperty -Name DCName -Value $RemoteComputerName
					$obj | Add-Member -MemberType NoteProperty -Name DCSite -Value $Site
					If($xIPAddress.Count -gt 1)
					{
						$obj | Add-Member -MemberType NoteProperty -Name DCIpAddress1 -Value $xIPAddress[0]
						$obj | Add-Member -MemberType NoteProperty -Name DCIpAddress2 -Value $xIPAddress[1]
					}
					Else
					{
						$obj | Add-Member -MemberType NoteProperty -Name DCIpAddress1 -Value $xIPAddress[0]
						$obj | Add-Member -MemberType NoteProperty -Name DCIpAddress2 -Value ""
					}
					
					$serverCount = If( $null -ne $nic.dnsserversearchorder ) { $nic.dnsserversearchorder.Length } Else { 0 }
					If( $serverCount -gt 0 )
					{
						$obj | Add-Member -MemberType NoteProperty -Name DCDNS1 -Value $xnicdnsserversearchorder[0]
						If( $serverCount -gt 1 )
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS2 -Value $xnicdnsserversearchorder[1]
						}
						Else
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS2 -Value " "
						}
						
						If( $serverCount -gt 2 )
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS3 -Value $xnicdnsserversearchorder[2]
						}
						Else
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS3 -Value " "
						}
						
						If( $serverCount -gt 3 )
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS4 -Value $xnicdnsserversearchorder[3]
						}
						Else
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS4 -Value " "
						}
					}

					$null = $Script:DCDNSIPInfo.Add( $obj )
				}
			}
		}	
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" -options $htmlbold
		}
	}
}
#endregion

#region word specific Functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. SMith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			#'de-'	{ 'Automatische Tabelle 2'; Break }
			'de-'	{ 'Automatisches Verzeichnis 2'; Break } #changed 6-feb-2022 rene bigler
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$ChineseArray -contains $_} {$CultureCode = "zh-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		
		If(($MSWord -eq $False) -and ($PDF -eq $True))
		{
			Write-Host "`n`n`t`tThis script uses Microsoft Word's SaveAs PDF function, please install Microsoft Word`n`n"
			AbortScript
		}
		Else
		{
			Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
			AbortScript
		}
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = $null –ne ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID})
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		AbortScript
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Set-DocumentProperty {
    <#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
    param (
        [object]$Document,
        [String]$DocProperty,
        [string]$Value
    )
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $builtInProperties = $Document.BuiltInDocumentProperties
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date -Format G): Setting up Word"
    
	If(!$AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName).pdf"
		}
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
		}
	}

	# Setup word for output
	Write-Verbose "$(Get-Date -Format G): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null

#Do not indent the following write-error lines. Doing so will mess up the console formatting of the error message.
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created. You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	The Word object could not be created. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	Write-Verbose "$(Get-Date -Format G): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	Unable to determine the Word language value. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}
	Write-Verbose "$(Get-Date -Format G): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	Microsoft Word 2007 is no longer supported.`n`n`t`tScript will end.
		`n`n
		"
		AbortScript
	}
	ElseIf($Script:WordVersion -eq 0)
	{
		Write-Error "
		`n`n
	The Word Version is 0. You should run a full online repair of your Office installation.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	You are running an untested or unsupported version of Microsoft Word.
		`n`n
	Script will end.
		`n`n
	Please send info on your version of Word to webster@carlwebster.com
		`n`n
		"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Verbose "$(Get-Date -Format G): Company name is blank. Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Host "
		Company Name is blank so Cover Page will not show a Company Name.
		Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value.
		You may want to use the -CompanyName parameter if you need a Company Name on the cover page.
			" -ForegroundColor White
			$Script:CoName = $TmpName
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date -Format G): Updated company name to $($Script:CoName)"
		}
	}
	Else
	{
		$Script:CoName = $CompanyName
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date -Format G): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date -Format G): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date -Format G): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date -Format G): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date -Format G): Culture code $($Script:WordCultureCode)"
		Write-Error "
		`n`n
	For $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.
		`n`n
	Script cannot Continue.
		`n`n
		"
		AbortScript
	}

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date -Format G): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object{$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date -Format G): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach-Object {
		If($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date -Format G): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Host "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist." -ForegroundColor White
		Write-Host "This report will not have a Cover Page." -ForegroundColor White
	}

	Write-Verbose "$(Get-Date -Format G): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An empty Word document could not be created. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An unknown error happened selecting the entire Word document for default formatting options.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 =.50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date -Format G): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date -Format G): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date -Format G): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date -Format G): "
			Write-Host "Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved." -ForegroundColor White
			Write-Host "This report will not have a Table of Contents." -ForegroundColor White
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Host "Table of Contents are not installed." -ForegroundColor White
		Write-Host "Table of Contents are not installed so this report will not have a Table of Contents." -ForegroundColor White
	}

	#set the footer
	Write-Verbose "$(Get-Date -Format G): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date -Format G): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date -Format G): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date -Format G): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#updated 8-Jun-2017 with additional cover page fields
	#Update document properties
	Write-Verbose "$(Get-Date -Format G): Set Cover Page Properties"
	#8-Jun-2017 put these 4 items in alpha order
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

	#Get the Coverpage XML part
	$cp = $Script:Doc.CustomXMLParts | Where-Object{$_.NamespaceURI -match "coverPageProps$"}

	#get the abstract XML part
	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "Abstract"}
	#set the text
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		[string]$abstract = $AbstractTitle
	}
	Else
	{
		[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
	}
	$ab.Text = $abstract

	#added 8-Jun-2017
	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyAddress"}
	#set the text
	[string]$abstract = $CompanyAddress
	$ab.Text = $abstract

	#added 8-Jun-2017
	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyEmail"}
	#set the text
	[string]$abstract = $CompanyEmail
	$ab.Text = $abstract

	#added 8-Jun-2017
	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyFax"}
	#set the text
	[string]$abstract = $CompanyFax
	$ab.Text = $abstract

	#added 8-Jun-2017
	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyPhone"}
	#set the text
	[string]$abstract = $CompanyPhone
	$ab.Text = $abstract

	$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "PublishDate"}
	#set the text
	[string]$abstract = (Get-Date -Format d).ToString()
	$ab.Text = $abstract

	Write-Verbose "$(Get-Date -Format G): Update the Table of Contents"
	#update the Table of Contents
	$Script:Doc.TablesOfContents.item(1).Update()
	$cp = $Null
	$ab = $Null
	$abstract = $Null
}
#endregion

#region registry Functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified local registry value or $Null if it is missing
Function Get-LocalRegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

Function Get-RegistryValue
{
	# Gets the specified registry value or $Null if it is missing
	[CmdletBinding()]
	Param
	(
		[String] $path, 
		[String] $name, 
		[String] $ComputerName
	)

	If($ComputerName -eq $env:computername -or $ComputerName -eq "LocalHost")
	{
		$key = Get-Item -LiteralPath $path -EA 0
		If($key)
		{
			Return $key.GetValue($name, $Null)
		}
		Else
		{
			Return $Null
		}
	}

	#path needed here is different for remote registry access
	$path1 = $path.SubString( 6 )
	$path2 = $path1.Replace( '\', '\\' )

	$registry = $null
	try
	{
		## use the Remote Registry service
		$registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
			[Microsoft.Win32.RegistryHive]::LocalMachine,
			$ComputerName ) 
	}
	catch
	{
		#$e = $error[ 0 ]
		#3.06, remove the verbose message as it confised some people
		#wv "Could not open registry on computer $ComputerName ($e)"
	}

	$val = $null
	If( $registry )
	{
		$key = $registry.OpenSubKey( $path2 )
		If( $key )
		{
			$val = $key.GetValue( $name )
			$key.Close()
		}

		$registry.Close()
	}

	Return $val
}
#endregion

#region word, text and html line output Functions
Function line
#Function created by Michael B. Smith, Exchange MVP
#@essentialexch on Twitter
#https://essential.exchange/blog
#for creating the formatted text report
#created March 2011
#updated March 2014
# updated March 2019 to use StringBuilder (about 100 times more efficient than simple strings)
{
	Param
	(
		[Int]    $tabs = 0, 
		[String] $name = '', 
		[String] $value = '', 
		[String] $newline = [System.Environment]::NewLine, 
		[Switch] $nonewline
	)

	while( $tabs -gt 0 )
	{
		#V3.00 - Switch to using a StringBuilder for $global:Output
		$null = $global:Output.Append( "`t" )
		$tabs--
	}

	If( $nonewline )
	{
		#V3.00 - Switch to using a StringBuilder for $global:Output
		$null = $global:Output.Append( $name + $value )
	}
	Else
	{
		#V3.00 - Switch to using a StringBuilder for $global:Output
		$null = $global:Output.AppendLine( $name + $value )
	}
}

## This will be replacing 'Line' - needs more testing
#Function created by Michael B. Smith, Exchange MVP
#@essentialexch on Twitter
#https://essential.exchange/blog
#for creating the formatted text report
#created March 2011
#updated March 2014
# updated March 2019 to use StringBuilder (about 100 times more efficient than simple strings)
# updated march 2019 to remove $newline param
# updated match 2019 to remove the $name and $value params
# -- now much smaller and much faster
Function lx
{
	Param
	(
		[Int]    $tabs = 0,
		[Switch] $nonewline
	)

	$null = $global:Output.Append( "`t" * $tabs )

	$str = $args -join ''

	If( $nonewline )
	{
		$null = $global:Output.Append( $str )
	}
	Else
	{
		$null = $global:Output.AppendLine( $str )
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This Function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlBold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlBold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlBold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML.  They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used.  Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

#V3.00
# to suppress $crlf in HTML documents, replace this with '' (empty string)
# but this was added to make the HTML readable
$crlf = [System.Environment]::NewLine

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
#errors with $HTMLStyle fixed 7-Dec-2017 by Webster
# re-implemented/re-based for v3.00 by Michael B. Smith
{
	Param
	(
		[Int]    $style    = 0, 
		[Int]    $tabs     = 0, 
		[String] $name     = '', 
		[String] $value    = '', 
		[String] $fontName = $null,
		[Int]    $fontSize = 1,
		[Int]    $options  = $htmlblack
	)

	#V3.00 - FIXME - long story short, this Function was wrong and had been wrong for a long time. 
	## The Function generated invalid HTML, and ignored fontname and fontsize parameters. I fixed
	## those items, but that made the report unreadable, because all of the formatting had been based
	## on this Function not working properly.

	## here is a typical H1 previously generated:
	## <h1>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\<font face='Calibri' color='#000000' size='1'></h1></font>

	## fixing the Function generated this (unreadably small):
	## <h1><font face='Calibri' color='#000000' size='1'>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\</font></h1>

	## So I took all the fixes out. This routine now generates valid HTML, but the fontName, fontSize,
	## and options parameters are ignored; so the routine generates equivalent output as before. I took
	## the fixes out instead of fixing all the call sites, because there are 225 call sites! If you are
	## willing to update all the call sites, you can easily re-instate the fixes. They have only been
	## commented out with '##' below.

	## If( [String]::IsNullOrEmpty( $fontName ) )
	## {
	##	$fontName = 'Calibri'
	## }
	## If( $fontSize -le 0 )
	## {
	##	$fontSize = 1
	## }

	## ## output data is stored here
	## [String] $output = ''
	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 1024 )

	If( [String]::IsNullOrEmpty( $name ) )	
	{
		## $HTMLBody = '<p></p>'
		$null = $sb.Append( '<p></p>' )
	}
	Else
	{
		## #V3.00
		[Bool] $ital  = $options -band $htmlitalics
		[Bool] $bold  = $options -band $htmlBold
		[Bool] $left  = $options -band $htmlAlignLeft
		[Bool] $right = $options -band $htmlAlignRight
		## $color = $global:htmlColor[ $options -band 0xfffff0 ]

		if( $left )
		{
			$HTMLBody += " align=left"
		}
		elseif( $right )
		{
			$HTMLBody += " align=right"
		}

		## ## build the HTML output string
##		$HTMLBody = ''
##		If( $ital ) { $HTMLBody += '<i>' }
##		If( $bold ) { $HTMLBody += '<b>' } 
		If( $ital ) { $null = $sb.Append( '<i>' ) }
		If( $bold ) { $null = $sb.Append( '<b>' ) } 

		Switch( $style )
		{
			1 { $HTMLOpen = '<h1'; $HTMLClose = '</h1>'; Break }
			2 { $HTMLOpen = '<h2'; $HTMLClose = '</h2>'; Break }
			3 { $HTMLOpen = '<h3'; $HTMLClose = '</h3>'; Break }
			4 { $HTMLOpen = '<h4'; $HTMLClose = '</h4>'; Break }
			Default { $HTMLOpen = ''; $HTMLClose = ''; Break }
		}

		## $HTMLBody += $HTMLOpen
		$null = $sb.Append( $HTMLOpen )
		if( $HTMLOpen.Length -gt 0 )
		{
			if( $left )
			{
				$null = $sb.Append( ' align=left' )
			}
			elseif( $right )
			{
				$null = $sb.Append( ' align=right' )
			}
			$null = $sb.Append( '>' )
		}

		## If($HTMLClose -eq '')
		## {
		##	$HTMLBody += "<br><font face='" + $fontName + "' " + "color='" + $color + "' size='"  + $fontSize + "'>"
		## }
		## Else
		## {
		##	$HTMLBody += "<font face='" + $fontName + "' " + "color='" + $color + "' size='"  + $fontSize + "'>"
		## }
		
##		while( $tabs -gt 0 )
##		{ 
##			$output += '&nbsp;&nbsp;&nbsp;&nbsp;'
##			$tabs--
##		}
		## output the rest of the parameters.
##		$output += $name + $value
		## $HTMLBody += $output
		$null = $sb.Append( ( '&nbsp;&nbsp;&nbsp;&nbsp;' * $tabs ) + $name + $value )

		## $HTMLBody += '</font>'
##		If( $HTMLClose -eq '' ) { $HTMLBody += '<br>'     }
##		Else                    { $HTMLBody += $HTMLClose }

##		If( $ital ) { $HTMLBody += '</i>' }
##		If( $bold ) { $HTMLBody += '</b>' } 

##		If( $HTMLClose -eq '' ) { $HTMLBody += '<br />' }

		If( $HTMLClose -eq '' ) { $null = $sb.Append( '<br>' )     }
		Else                    { $null = $sb.Append( $HTMLClose ) }

		If( $ital ) { $null = $sb.Append( '</i>' ) }
		If( $bold ) { $null = $sb.Append( '</b>' ) } 

		If( $HTMLClose -eq '' ) { $null = $sb.Append( '<br />' ) }
	}
	##$HTMLBody += $crlf
	$null = $sb.AppendLine( '' )

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $sb.ToString() 4>$Null
}
#endregion

#region HTML table Functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable Function
# Created by Ken Avram
# modified by Jake Rutski
# re-implemented by Michael B. Smith for v3.00. Also made the documentation match reality.
#***********************************************************************************************************
Function AddHTMLTable
{
	Param
	(
		[String]   $fontName  = 'Calibri',
		[Int]      $fontSize  = 2,
		[Int]      $colCount  = 0,
		[Int]      $rowCount  = 0,
		[Object[]] $rowInfo   = $null,
		[Object[]] $fixedInfo = $null
	)
	#V3.00 - Use StringBuilder - MBS
	## In the normal case, tables are only a few dozen cells. But in the case
	## of Sites, OUs, and Users - there may be many hundreds of thousands of 
	## cells. Using normal strings is too slow.

	#V3.00
	## If( $ExtraSpecialVerbose )
	## {
	##	$global:rowInfo1 = $rowInfo
	## }
<#
	If( $SuperVerbose )
	{
		wv "AddHTMLTable: fontName '$fontName', fontsize $fontSize, colCount $colCount, rowCount $rowCount"
		If( $null -ne $rowInfo -and $rowInfo.Count -gt 0 )
		{
			wv "AddHTMLTable: rowInfo has $( $rowInfo.Count ) elements"
			If( $ExtraSpecialVerbose )
			{
				wv "AddHTMLTable: rowInfo length $( $rowInfo.Length )"
				for( $ii = 0; $ii -lt $rowInfo.Length; $ii++ )
				{
					$row = $rowInfo[ $ii ]
					wv "AddHTMLTable: index $ii, type $( $row.GetType().FullName ), length $( $row.Length )"
					for( $yyy = 0; $yyy -lt $row.Length; $yyy++ )
					{
						wv "AddHTMLTable: index $ii, yyy = $yyy, val = '$( $row[ $yyy ] )'"
					}
					wv "AddHTMLTable: done"
				}
			}
		}
		Else
		{
			wv "AddHTMLTable: rowInfo is empty"
		}
		If( $null -ne $fixedInfo -and $fixedInfo.Count -gt 0 )
		{
			wv "AddHTMLTable: fixedInfo has $( $fixedInfo.Count ) elements"
		}
		Else
		{
			wv "AddHTMLTable: fixedInfo is empty"
		}
	}
#>

	$fwLength = if( $null -ne $fixedInfo ) { $fixedInfo.Count } else { 0 }

	##$htmlbody = ''
	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 8192 )

	If( $rowInfo -and $rowInfo.Length -lt $rowCount )
	{
##		$oldCount = $rowCount
		$rowCount = $rowInfo.Length
##		If( $SuperVerbose )
##		{
##			wv "AddHTMLTable: updated rowCount to $rowCount from $oldCount, based on rowInfo.Length"
##		}
	}

	for( $rowCountIndex = 0; $rowCountIndex -lt $rowCount; $rowCountIndex++ )
	{
		$null = $sb.AppendLine( '<tr>' )
		## $htmlbody += '<tr>'
		## $htmlbody += $crlf #V3.00 - make the HTML readable

		## each row of rowInfo is an array
		## each row consists of tuples: an item of text followed by an item of formatting data
<#		
		$row = $rowInfo[ $rowCountIndex ]
		If( $ExtraSpecialVerbose )
		{
			wv "!!!!! AddHTMLTable: rowCountIndex = $rowCountIndex, row.Length = $( $row.Length ), row gettype = $( $row.GetType().FullName )"
			wv "!!!!! AddHTMLTable: colCount $colCount"
			wv "!!!!! AddHTMLTable: row[0].Length $( $row[0].Length )"
			wv "!!!!! AddHTMLTable: row[0].GetType $( $row[0].GetType().FullName )"
			$subRow = $row
			If( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
			{
				$subRow = $subRow[ 0 ]
				wv "!!!!! AddHTMLTable: deref subRow.Length $( $subRow.Length ), subRow.GetType $( $subRow.GetType().FullName )"
			}

			for( $columnIndex = 0; $columnIndex -lt $subRow.Length; $columnIndex += 2 )
			{
				$item = $subRow[ $columnIndex ]
				wv "!!!!! AddHTMLTable: item.GetType $( $item.GetType().FullName )"
				## If( !( $item -is [String] ) -and $item -is [Array] )
##				If( $item -is [Array] -and $item[ 0 ] -is [Array] )				
##				{
##					$item = $item[ 0 ]
##					wv "!!!!! AddHTMLTable: dereferenced item.GetType $( $item.GetType().FullName )"
##				}
				wv "!!!!! AddHTMLTable: rowCountIndex = $rowCountIndex, columnIndex = $columnIndex, val '$item'"
			}
			wv "!!!!! AddHTMLTable: done"
		}
#>

		## reset
		$row = $rowInfo[ $rowCountIndex ]

		$subRow = $row
		If( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
		{
			$subRow = $subRow[ 0 ]
			## wv "***** AddHTMLTable: deref rowCountIndex $rowCountIndex, subRow.Length $( $subRow.Length ), subRow.GetType $( $subRow.GetType().FullName )"
		}

		$subRowLength = $subRow.Count
		for( $columnIndex = 0; $columnIndex -lt $colCount; $columnIndex += 2 )
		{
			$item = If( $columnIndex -lt $subRowLength ) { $subRow[ $columnIndex ] } Else { 0 }
			## If( !( $item -is [String] ) -and $item -is [Array] )
##			If( $item -is [Array] -and $item[ 0 ] -is [Array] )
##			{
##				$item = $item[ 0 ]
##			}

			$text   = If( $item ) { $item.ToString() } Else { '' }
			$format = If( ( $columnIndex + 1 ) -lt $subRowLength ) { $subRow[ $columnIndex + 1 ] } Else { 0 }
			## item, text, and format ALWAYS have values, even if empty values
			$color  = $global:htmlColor[ $format -band 0xfffff0 ]
			[Bool] $bold  = $format -band $htmlBold
			[Bool] $ital  = $format -band $htmlitalics
			[Bool] $left  = $format -band $htmlAlignLeft
			[Bool] $right = $format -band $htmlAlignRight		

<#			
			If( $ExtraSpecialVerbose )
			{
				wv "***** columnIndex $columnIndex, subRow.Length $( $subRow.Length ), item GetType $( $item.GetType().FullName ), item '$item'"
				wv "***** format $format, color $color, text '$text'"
				wv "***** format gettype $( $format.GetType().Fullname ), text gettype $( $text.GetType().Fullname )"
			}
#>

			If( $fwLength -eq 0 )
			{
				$null = $sb.Append( "<td style=""background-color:$( $color )""" )
				##$htmlbody += "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>"
			}
			Else
			{
				$null = $sb.Append( "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""" )
				##$htmlbody += "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>"
			}
			if( $left )
			{
				$null = $sb.Append( ' align=left' )
			}
			elseif( $right )
			{
				$null = $sb.Append( ' align=right' )
			}
			$null = $sb.Append( "><font face='$($fontName)' size='$($fontSize)'>" )

			##If( $bold ) { $htmlbody += '<b>' }
			##If( $ital ) { $htmlbody += '<i>' }
			If( $bold ) { $null = $sb.Append( '<b>' ) }
			If( $ital ) { $null = $sb.Append( '<i>' ) }

			If( $text -eq ' ' -or $text.length -eq 0)
			{
				##$htmlbody += '&nbsp;&nbsp;&nbsp;'
				$null = $sb.Append( '&nbsp;&nbsp;&nbsp;' )
			}
			Else
			{
				for ($inx = 0; $inx -lt $text.length; $inx++ )
				{
					If( $text[ $inx ] -eq ' ' )
					{
						##$htmlbody += '&nbsp;'
						$null = $sb.Append( '&nbsp;' )
					}
					Else
					{
						break
					}
				}
				##$htmlbody += $text
				$null = $sb.Append( $text )
			}

##			If( $bold ) { $htmlbody += '</b>' }
##			If( $ital ) { $htmlbody += '</i>' }
			If( $bold ) { $null = $sb.Append( '</b>' ) }
			If( $ital ) { $null = $sb.Append( '</i>' ) }

			$null = $sb.AppendLine( '</font></td>' )
##			$htmlbody += '</font></td>'
##			$htmlbody += $crlf
		}

		$null = $sb.AppendLine( '</tr>' )
##		$htmlbody += '</tr>'
##		$htmlbody += $crlf
	}

##	If( $ExtraSpecialVerbose )
##	{
##		$global:rowInfo = $rowInfo
##		wv "!!!!! AddHTMLTable: HTML = '$htmlbody'"
##	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $sb.ToString() 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
# reworked by Michael B. Smith for v3.00
#***********************************************************************************************************

<#
.Synopsis
	Formats table column headers for an HTML table.
.DESCRIPTION
	This function formats table column headers for an HTML table. It requires 
	AddHTMLTable to format the individual rows of the table.

.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border = '0'). 
	Otherwise the table will be generated with a border (border = '1').

.PARAMETER noHeadCols
	This parameter should be used when generating tables which do not have a 
	separate array containing column headers (columnArray is not specified). 

	Set this parameter equal to the number of (header) columns in the table.

.PARAMETER rowArray
	This parameter contains the row data array for the table.

	The total numbers of rows in the table is equal to $rowArray.Length + $tableHeader.Length.
	$tableHeader.Length may be zero (the parameter can be $null).

	Each entry in rowarray is ANOTHER array of tuples. The first element of the tuple is the 
	contents of the cell, and the second element of the tuple is the color of the cell, then
	they duplicate for every cell in the row.

.PARAMETER columnArray
	This parameter contains column header data for the table.

	The total number of columns in the table is equal to $columnarray.Length or $null.

	If $columnarray is $null, then there are no column headers, just the first line of the
	table and noHeadCols is used to size the table.

	Each entry in $columnarray organized as a set of two items. The first is the
	data for the header cell. THe second is the color/italic/bold for the header cell. So
	the total number of columns is ($columnArray.Length / 2) when $columnArray isn't $null.

	I have no idea why it wasn't done identically to $rowarray.

.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $fixedWidth = @("100px","110px","120px","130px","140px")

	This is mapped to both rowArray and columnArray.

.PARAMETER tableHeader
	A string containing the header for the table (printed at the top of the table, left justified). The
	default is a blank string.
.PARAMETER tableWidth
	The width of the table in pixels, or 'auto'. The default is 'auto'.
.PARAMETER fontName
	The name of the font to use in the table. The default is 'Calibri'.
.PARAMETER fontSize
	The size of the font to use in the table. The default is 2. Note that this is the HTML size, not the pixel size.

.USAGE
	FormatHTMLTable <Table Header> <Table Width> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file.  All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column.  You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column.  Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data.  Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array.  If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',$htmlsb,'Status',$htmlsb,'Startup Type',$htmlsb)

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics.  For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below.  As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",$htmlsb,$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',$htmlsb,$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',$htmlsb,$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',$htmlsb,$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',$htmlsb,$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',$htmlsb,$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',$htmlsb,$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',$htmlsb,$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',$htmlsb,$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',$htmlsb,$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',$htmlsb,$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',$htmlsb,$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the Function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param
	(
		[String]   $tableheader = '',
		[String]   $tablewidth  = 'auto',
		[String]   $fontName    = 'Calibri',
		[Int]      $fontSize    = 2,
		[Switch]   $noBorder    = $false,
		[Int]      $noHeadCols  = 1,
		[Object[]] $rowArray    = $null,
		[Object[]] $fixedWidth  = $null,
		[Object[]] $columnArray = $null
	)

	## FIXME - the help text for this Function is wacky wrong - MBS
	## FIXME - Use StringBuilder - MBS - this only builds the table header - benefit relatively small
<#
	If( $SuperVerbose )
	{
		wv "FormatHTMLTable: fontname '$fontname', size $fontSize, tableheader '$tableheader'"
		wv "FormatHTMLTable: noborder $noborder, noheadcols $noheadcols"
		If( $rowarray -and $rowarray.count -gt 0 )
		{
			wv "FormatHTMLTable: rowarray has $( $rowarray.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: rowarray is empty"
		}
		If( $columnarray -and $columnarray.count -gt 0 )
		{
			wv "FormatHTMLTable: columnarray has $( $columnarray.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: columnarray is empty"
		}
		If( $fixedwidth -and $fixedwidth.count -gt 0 )
		{
			wv "FormatHTMLTable: fixedwidth has $( $fixedwidth.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: fixedwidth is empty"
		}
	}
#>

	$HTMLBody = ''
	if( $tableheader.Length -gt 0 )
	{
		$HTMLBody += "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>" + $crlf
	}

	$fwSize = if( $null -eq $fixedWidth ) { 0 } else { $fixedWidth.Count }

	If( $null -eq $columnArray -or $columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If( $null -eq $rowArray )
	{
		$NumRows = 1
	}
	Else
	{
		$NumRows = $rowArray.length + 1
	}

	If( $noBorder )
	{
		$HTMLBody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$HTMLBody += "<table border='1' width='" + $tablewidth + "'>"
	}
	$HTMLBody += $crlf

	If( $columnArray -and $columnArray.Length -gt 0 )
	{
		$HTMLBody += '<tr>' + $crlf

		for( $columnIndex = 0; $columnIndex -lt $NumCols; $columnindex += 2 )
		{
			#V3.00
			$val = $columnArray[ $columnIndex + 1 ]
			$tmp = $global:htmlColor[ $val -band 0xfffff0 ]
			[Bool] $bold  = $val -band $htmlBold
			[Bool] $ital  = $val -band $htmlitalics
			[Bool] $left  = $val -band $htmlAlignLeft
			[Bool] $right = $val -band $htmlAlignRight		

			If( $fwSize -eq 0 )
			{
				$HTMLBody += "<td style=""background-color:$($tmp)"""
			}
			Else
			{
				$HTMLBody += "<td style=""width:$($fixedWidth[$columnIndex / 2]); background-color:$($tmp)"""
			}
			if( $left )
			{
				$HTMLBody += " align=left"
			}
			elseif( $right )
			{
				$HTMLBody += " align=right"
			}

			$HTMLBody += "><font face='$($fontName)' size='$($fontSize)'>"

			If( $bold ) { $HTMLBody += '<b>' }
			If( $ital ) { $HTMLBody += '<i>' }

			$array = $columnArray[ $columnIndex ]
			If( $array )
			{
				If( $array -eq ' ' -or $array.Length -eq 0 )
				{
					$HTMLBody += '&nbsp;&nbsp;&nbsp;'
				}
				Else
				{
					for( $i = 0; $i -lt $array.Length; $i += 2 )
					{
						If( $array[ $i ] -eq ' ' )
						{
							$HTMLBody += '&nbsp;'
						}
						Else
						{
							break
						}
					}
					$HTMLBody += $array
				}
			}
			Else
			{
				$HTMLBody += '&nbsp;&nbsp;&nbsp;'
			}
			
			If( $bold ) { $HTMLBody += '</b>' }
			If( $ital ) { $HTMLBody += '</i>' }

			$HTMLBody += '</font></td>'
			$HTMLBody += $crlf
		}

		$HTMLBody += '</tr>' + $crlf
	}

	#V3.00
	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $HTMLBody 4>$Null 
	$HTMLBody = ''

	##$rowindex = 2
	If( $rowArray )
	{
<#
		If( $ExtraSpecialVerbose )
		{
			wv "***** FormatHTMLTable: rowarray length $( $rowArray.Length )"
			for( $ii = 0; $ii -lt $rowArray.Length; $ii++ )
			{
				$row = $rowArray[ $ii ]
				wv "***** FormatHTMLTable: index $ii, type $( $row.GetType().FullName ), length $( $row.Length )"
				for( $yyy = 0; $yyy -lt $row.Length; $yyy++ )
				{
					wv "***** FormatHTMLTable: index $ii, yyy = $yyy, val = '$( $row[ $yyy ] )'"
				}
				wv "***** done"
			}
			wv "***** FormatHTMLTable: rowCount $NumRows"
		}
#>

		AddHTMLTable -fontName $fontName -fontSize $fontSize `
			-colCount $numCols -rowCount $NumRows `
			-rowInfo $rowArray -fixedInfo $fixedWidth
		##$rowArray = @()
		$rowArray = $null
		$HTMLBody = '</table>'
	}
	Else
	{
		$HTMLBody += '</table>'
	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML Functions
<#
#***********************************************************************************************************
# CheckHTMLColor - Called from AddHTMLTable WriteHTMLLine and FormatHTMLTable
#***********************************************************************************************************
Function CheckHTMLColor
{
	Param($hash)

	#V3.00 -- this is really slow. several ways to fixit. so fixit. MBS
	#V3.00 - obsolete. replaced by using $global:htmlColor lookup table
	If($hash -band $htmlwhite)
	{
		Return $htmlwhitemask
	}
	If($hash -band $htmlred)
	{
		Return $htmlredmask
	}
	If($hash -band $htmlcyan)
	{
		Return $htmlcyanmask
	}
	If($hash -band $htmlblue)
	{
		Return $htmlbluemask
	}
	If($hash -band $htmldarkblue)
	{
		Return $htmldarkbluemask
	}
	If($hash -band $htmllightblue)
	{
		Return $htmllightbluemask
	}
	If($hash -band $htmlpurple)
	{
		Return $htmlpurplemask
	}
	If($hash -band $htmlyellow)
	{
		Return $htmlyellowmask
	}
	If($hash -band $htmllime)
	{
		Return $htmllimemask
	}
	If($hash -band $htmlmagenta)
	{
		Return $htmlmagentamask
	}
	If($hash -band $htmlsilver)
	{
		Return $htmlsilvermask
	}
	If($hash -band $htmlgray)
	{
		Return $htmlgraymask
	}
	If($hash -band $htmlblack)
	{
		Return $htmlblackmask
	}
	If($hash -band $htmlorange)
	{
		Return $htmlorangemask
	}
	If($hash -band $htmlmaroon)
	{
		Return $htmlmaroonmask
	}
	If($hash -band $htmlgreen)
	{
		Return $htmlgreenmask
	}
	If($hash -band $htmlolive)
	{
		Return $htmlolivemask
	}
}
#>

Function SetupHTML
{
	Write-Verbose "$(Get-Date -Format G): Setting up HTML"
	If(!$AddDateTime)
	{
		[string]$Script:HTMLFileName = "$($Script:pwdpath)\$($OutputFileName).html"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:HTMLFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	out-file -FilePath $Script:HTMLFileName -Force -InputObject $HTMLHead 4>$Null
}
#endregion

#region Iain's Word table Functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This Function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this Function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is Returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -ne $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end Switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This Function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells Returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells Returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) Returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] [int]$BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end Switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This Function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This Function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this Function is called by the AddWordTable Function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this Functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script Functions
Function Get-ADTrustInfo
{
	Param ($TrustInfo)

	$propertyhash = @{
		TrustType = $NULL
		TrustAttribute = $NULL
		TrustDirection = $NULL
	}

	$TrustDirectionNumber = $TrustInfo.TrustDirection
	$TrustTypeNumber = $TrustInfo.TrustType
	$TrustAttributesNumber = $TrustInfo.TrustAttributes
	
	#http://msdn.microsoft.com/en-us/library/cc234293.aspx
	Switch ($TrustTypeNumber) 
	{ 
		1 { $propertyhash['TrustType'] = "Trust with a Windows domain not running Active Directory"; Break} 
		2 { $propertyhash['TrustType'] = "Trust with a Windows domain running Active Directory"; Break} 
		3 { $propertyhash['TrustType'] = "Trust with a non-Windows-compliant Kerberos distribution"; Break} 
		4 { $propertyhash['TrustType'] = "Trust with a DCE realm (not used)"; Break} 
		Default { $propertyhash['TrustType'] = "Invalid Trust Type of $($TrustTypeNumber)" ; Break}
	}
	
	$propertyhash['TrustAttribute'] = @()
	#$hextrustAttributesValue = '{0:X}' -f $trustAttributesNumber
	Switch ($trustAttributesNumber)
	{
		1 {$propertyhash['TrustAttribute'] += "Non-Transitive"}
		2 {$propertyhash['TrustAttribute'] += "Uplevel clients only"}
		4 {$propertyhash['TrustAttribute'] += "Quarantined Domain (External, SID Filtering)"}
		8 {$propertyhash['TrustAttribute'] += "Cross-Organizational Trust (Selective Authentication)"}
		16 {$propertyhash['TrustAttribute'] += "Interforest Trust"}
		32 {$propertyhash['TrustAttribute'] += "Intraforest Trust"}
		64 {$propertyhash['TrustAttribute'] += "MIT Trust using RC4 Encryption"}
		512 {$propertyhash['TrustAttribute'] += "Cross organization Trust no TGT delegation"}
	}
	
	Switch ($TrustDirectionNumber) 
	{ 
		0 { $propertyhash['TrustDirection'] = "Disabled"; Break} 
		1 { $propertyhash['TrustDirection'] = "Inbound"; Break} 
		2 { $propertyhash['TrustDirection'] = "Outbound"; Break} 
		3 { $propertyhash['TrustDirection'] = "Bidirectional"; Break} 
		Default { $propertyhash['TrustDirection'] = $TrustDirectionNumber ; Break}
	}
	
	New-Object -Type PSObject -property $propertyhash
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#Function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			If((Get-Member -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function validObject( [object] $object, [string] $topLevel )
{
	#Function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			Return $True
		}
	}
	Return $False
}

Function BuildMultiColumnTable
{
	Param([Array]$xArray)
	
	#divide by 0 bug reported 9-Apr-2014 by Lee Dehmer 
	#if security group name or OU name was longer than 60 characters it caused a divide by 0 error
	
	#added a second parameter to the Function so the verbose message would say whether 
	#the Function is processing servers, security groups or OUs.
	#V3.03, remove the second parameter as it was no longer needed or used
	
	If(-not ($xArray -is [Array]))
	{
		$xArray = (,$xArray)
	}
	
	[int]$MaxLength = 0
	[int]$TmpLength = 0
	#remove 60 as a hard-coded value
	#60 is the max width the table can be when indented 36 points
	[int]$MaxTableWidth = 60
	ForEach($xName in $xArray)
	{
		$TmpLength = $xName.Length
		If($TmpLength -gt $MaxLength)
		{
			$MaxLength = $TmpLength
		}
	}
	$TableRange = $doc.Application.Selection.Range
	#removed hard-coded value of 60 and replace with MaxTableWidth variable
	#fix in 3.03 if $MaxLength was 0, prevent division by 0
	If($MaxLength -eq 0)
	{
		[int]$Columns = 0
	}
	Else
	{
		[int]$Columns = [Math]::Floor($MaxTableWidth / $MaxLength)
	}
	
	If($xArray.count -lt $Columns)
	{
		[int]$Rows = 1
		#not enough array items to fill columns so use array count
		$MaxCells  = $xArray.Count
		#reset column count so there are no empty columns
		$Columns   = $xArray.Count 
	}
	ElseIf($Columns -eq 0)
	{
		#divide by 0 bug if this condition is not handled
		#number was larger than $MaxTableWidth so there can only be one column
		#with one cell per row
		[int]$Rows = $xArray.count
		$Columns   = 1
		$MaxCells  = 1
	}
	Else
	{
		[int]$Rows = [Math]::Floor( ( $xArray.count + $Columns - 1 ) / $Columns)
		#more array items than columns so don't go past last column
		$MaxCells  = $Columns
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$Table.Style = $Script:MyHash.Word_TableGrid
	
	$Table.Borders.InsideLineStyle = $wdLineStyleSingle
	$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
	[int]$xRow = 1
	[int]$ArrayItem = 0
	While($xRow -le $Rows)
	{
		For($xCell=1; $xCell -le $MaxCells; $xCell++)
		{
			$Table.Cell($xRow,$xCell).Range.Text = $xArray[$ArrayItem]
			$ArrayItem++
			If($ArrayItem -eq $xArray.Count) #fix in 3.03 to catch a weird array out of bounds error
			{
				Break
			}
		}
		$xRow++
	}
	$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
	$Table.AutoFitBehavior($wdAutoFitContent)
	
	FindWordDocumentEnd
	$TableRange = $Null
	$Table = $Null
	$xArray = $Null
}

Function UserIsaDomainAdmin
{
	#Function adapted from sample code provided by Thomas Vuylsteke
	$IsDA = $False
	$name = $env:username
	Write-Verbose "$(Get-Date -Format G): `t`tTokenGroups - Checking groups for $name"

	$root = [ADSI]""
	$filter = "(sAMAccountName=$name)"
	$props = @("distinguishedName")
	$Searcher = new-Object System.DirectoryServices.DirectorySearcher($root,$filter,$props)
	$account = $Searcher.FindOne().properties.distinguishedname

	$user = [ADSI]"LDAP://$Account"
	$user.GetInfoEx(@("tokengroups"),0)
	$groups = $user.Get("tokengroups")

	$domainAdminsSID = New-Object System.Security.Principal.SecurityIdentifier (((Get-ADDomain -Server $ADForest -EA 0).DomainSid).Value+"-512") 

	ForEach($group in $groups)
	{     
		$ID = New-Object System.Security.Principal.SecurityIdentifier($group,0)       
		If($ID.CompareTo($domainAdminsSID) -eq 0)
		{
			$IsDA = $True
			Break
		}     
	}

	Return $IsDA
}

Function ElevatedSession
{
	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

	If($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
	{
		Write-Verbose "$(Get-Date -Format G): This is an elevated PowerShell session"
		Return $True
	}
	Else
	{
		Write-Host "" -Foreground White
		Write-Host "$(Get-Date): This is NOT an elevated PowerShell session" -Foreground White
		Write-Host "" -Foreground White
		Return $False
	}
}

Function Get-ComputerCountByOS
{
	<#
	This Function is provided by Jeremy Saunders and used with his permission

	http://www.jhouseconsulting.com/2014/06/22/script-to-create-an-overview-of-all-computer-objects-in-a-domain-1385

	Jeremy sent me version 1.8 of his script to use as the basis for this Function
	
	This Function will provide an overview and count of all computer objects in a
	domain based on Operating System and Service Pack. It helps an organisation
	to understand the number of stale and active computers against the different
	types of operating systems deployed in their environment.

	Computer objects are filtered into 4 categories:
	1) Windows Servers
	2) Windows Workstations
	3) Other non-Windows (Linux, Mac, etc)
	4) Windows Cluster Name Objects (CNOs) and Virtual Computer Objects (VCOs)

	A Stale object is derived from 2 values ANDed together:
	1) PasswordLastChanged > $MaxPasswordLastChanged days ago
	2) LastLogonDate > $MaxLastLogonDate days ago

	By default the script variable for $MaxPasswordLastChanged is set to 90 and
	the variable for $MaxLastLogonDate is set to 30. These can easily be adjusted
	to suite your definition of a stale object.

	The Active objects column is calculated by subtracting the Enabled_Stale
	value from the Enabled value. This gives us an accurate number of active
	objects against each Operating System.

	To help provide a high level overview of the computer object landscape, we
	calculate the number of stale objects of enabled and disabled objects
	separately. Disabled objects are often ignored, but it's pointless leaving
	old disabled computer objects in the domain.

	For viewing purposes we sort the output by Operating System and not count.

	You may notice a question mark (?) in some of the OperatingSystem strings.
	This is a representation of each Double-Byte character that was unable to
	be translated. Refer to Microsoft KB829856 for an explanation.

	Computers change their passwword if and when they feel like it. The domain
	doesn't initiate the change. It is controlled by three values under the
	following registry key:
	- HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters
	- DisablePasswordChange
	- MaximumPasswordAge
	- RefusePasswordChange
	If these values are not present, the default value of 30 days will be used.

	Non Windows operating systems and appliances vary in the way they manage
	their password change, with some not doing it at all. It's best to do some
	research before drawing any conclusions about the validity of these objects.
	i.e. Don't just go and delete them because my script says they are stale.

	Be aware that a cluster updates the lastLogonTimeStamp of the CNO/VNO when
	it brings a clustered network name resource online. So it could be running
	for months without an update to the lastLogonTimeStamp attribute.

	Syntax examples:

	- To execute the script in the current Domain:
	Get-ComputerCountByOS.ps1

	- To execute the script against a trusted Domain:
	Get-ComputerCountByOS.ps1 -TrustedDomain mydemosthatrock.com

	Script Name: Get-ComputerCountByOS.ps1
	Release: 1.6
	Written by Jeremy@jhouseconsulting.com 20th May 2012
	Modified by Jeremy@jhouseconsulting.com 4th December 2015

	#>

	#-------------------------------------------------------------
	Param
	(
		[String] $TrustedDomain
	)

	#-------------------------------------------------------------

	# Set this to true to include service pack level. This makes the
	# output more granular, as the counts are then based on Operating
	# System + Service Pack.
	$OperatingSystemIncludesServicePack = $True

	# Set this to the maximum value in number of days when the computer
	# password last changed. Do not go beyond 90 days.
	$MaxPasswordLastChanged = 90

	# Set this to the maximum value in number of days when the computer
	# last logged onto the domain.
	$MaxLastLogonDate = 30

	#-------------------------------------------------------------

	$TotalComputersProcessed   = 0
	$ComputerCount             = 0
	$TotalStaleObjects         = 0
	$TotalEnabledStaleObjects  = 0
	$TotalEnabledObjects       = 0
	$TotalDisabledObjects      = 0
	$TotalDisabledStaleObjects = 0
	$AllComputerObjects        = New-Object System.Collections.ArrayList
	$WindowsServerObjects      = New-Object System.Collections.ArrayList
	$WindowsWorkstationObjects = New-Object System.Collections.ArrayList
	$NonWindowsComputerObjects = New-Object System.Collections.ArrayList
	$CNOandVCOObjects          = New-Object System.Collections.ArrayList
	$ComputersHashTable        = @{}

	$context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext( "domain", $TrustedDomain )
	Try 
	{
		$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain( $context )
	}
	Catch [exception] 
	{
		Write-Error $_.Exception.Message
		Exit
	}

	# Get AD Distinguished Name
	$DomainDistinguishedName = $Domain.GetDirectoryEntry().DistinguishedName.Value

	$ADSearchBase = $DomainDistinguishedName

	Write-Verbose "$(Get-Date -Format G): `t`tGathering computer misc data"

	# Create an LDAP search for all computer objects
	$ADFilter = '(objectCategory=computer)'

	# There is a known bug in PowerShell requiring the DirectorySearcher
	# properties to be in lower case for reliability.
	$ADPropertyList = @( "distinguishedname", "name", "operatingsystem", "operatingsystemversion",
		"operatingsystemservicepack", "description", "info", "useraccountcontrol",
		"pwdlastset", "lastlogontimestamp", "whencreated", "serviceprincipalname" )

	$ADScope                = 'subtree'
	$ADPageSize             = 1000
	$ADSearchRoot           = New-Object System.DirectoryServices.DirectoryEntry( "LDAP://$($ADSearchBase)" ) 
	$ADSearcher             = New-Object System.DirectoryServices.DirectorySearcher
	$ADSearcher.SearchRoot  = $ADSearchRoot
	$ADSearcher.PageSize    = $ADPageSize 
	$ADSearcher.Filter      = $ADFilter 
	$ADSearcher.SearchScope = $ADScope
	If( $ADPropertyList ) 
	{
		ForEach( $ADProperty in $ADPropertyList ) 
		{
			$null = $ADSearcher.PropertiesToLoad.Add( $ADProperty )
		}
	}

	Try 
	{
		Write-Verbose "Please be patient whilst the script retrieves all computer objects and specified attributes..."
		$colResults = $ADSearcher.Findall()
		# Dispose of the search and results properly to avoid a memory leak
		$ADSearcher.Dispose()
		$ComputerCount = $colResults.Count
	}
	Catch 
	{
		$ComputerCount = 0
		Write-Warning "The $ADSearchBase structure cannot be found!"
	}

	If($ComputerCount -ne 0) 
	{
		Write-Verbose "Processing $ComputerCount computer objects in the $domain Domain..."
		ForEach($objResult in $colResults) 
		{
			$Name = $objResult.Properties[ 'name' ].Item( 0 )
			$DistinguishedName = $objResult.Properties[ 'distinguishedname' ].Item( 0 )

			$ParentDN = $DistinguishedName -split '(?<![\\]),' ## this is a lot of trouble to go through to allow a
			                                                   ## computer name with an escaped ',' -- which isn't legal
			$ParentDN = $ParentDN[1..$($ParentDN.Count-1)] -join ','

			## using Properties as an array allows eliminating all the prior
			## expensive try/catch blocks
			$val = $objResult.Properties[ 'operatingsystem' ]
			If( $null -ne $val -and $val.Count -gt 0 )
			{
				$OperatingSystem = $val.Item( 0 )
			} 
			Else 
			{
				$OperatingSystem = 'Undefined'
			}

			$val = $objResult.Properties[ 'operatingsystemversion' ]
			If( $null -ne $val -and $val.Count -gt 0 ) 
			{
				$OperatingSystemVersion = $val.Item( 0 )
			} 
			Else 
			{
				$OperatingSystemVersion = ''
			}

			$val = $objResult.Properties[ 'operatingsystemservicepack' ]
			If( $null -ne $val -and $val.Count -gt 0 )
			{					
				$OperatingSystemServicePack = $val.Item( 0 )
			}
			Else 
			{
				$OperatingSystemServicePack = ''
			}

			$val = $objResult.Properties[ 'description' ]
			If( $null -ne $val -and $val.Count -gt 0 ) 
			{
				$Description = $val.Item( 0 )
			} 
			Else 
			{
				$Description = ''
			}

			$PasswordTooOld = $False
			$PasswordLastSet = [System.DateTime]::FromFileTime( $objResult.Properties[ 'pwdlastset' ].Item( 0 ) )
			If($PasswordLastSet -lt (Get-Date).AddDays(-$MaxPasswordLastChanged)) 
			{
				$PasswordTooOld = $True
			}

			$HasNotRecentlyLoggedOn = $False

			$val = $objResult.Properties[ 'lastlogontimestamp' ]
			If( $null -ne $val -and $val.Count -gt 0 )
			{
				$LastLogonTimeStamp = $val.Item( 0 )
				$LastLogon = [System.DateTime]::FromFileTime( $LastLogonTimeStamp )
				If( $LastLogon -le (Get-Date).AddDays( -$MaxLastLogonDate ) ) 
				{
					$HasNotRecentlyLoggedOn = $True
				}
				#If( $LastLogon -match "1/01/1601" )  ## FIXME - is this accurate for all cultures???
				If($LastLogonTimeStamp -eq 0) #changed in V3.00
				{
					$LastLogon = 'Never logged on before'
				}
			} 
			Else 
			{
				$LastLogon = 'Never logged on before'
			}

			$WhenCreated = $objResult.Properties[ 'whencreated' ].Item( 0 )

			# If it's never logged on before and was created more than $MaxLastLogonDate days
			# ago, set the $HasNotRecentlyLoggedOn variable to True.
			# An example of this would be if you prestaged the account but never ended up using
			# it.
			If( $lastLogon -eq 'Never logged on before' ) 
			{
				If( $whencreated -le ( Get-Date ).AddDays( -$MaxLastLogonDate ) ) 
				{
					$HasNotRecentlyLoggedOn = $True
				}
			}

			# Check if it's a stale object.
			$IsStale = $False
			If($PasswordTooOld -eq $True -AND $HasNotRecentlyLoggedOn -eq $True) 
			{
				$IsStale = $True
			}

			$val = $objResult.Properties[ 'serviceprincipalname' ]
			If( $null -ne $val -and $val.Count -gt 0 )
			{
				$ServicePrincipalName = $val.Item( 0 )
			}
			Else
			{
				$ServicePrincipalName = ''
			}

			$UserAccountControl = $objResult.Properties[ 'useraccountcontrol' ].Item( 0 )
			$Enabled = $True
			If( ( $UserAccountControl -bor 0x0002 ) -eq $UserAccountControl )
			{
				$Enabled = $False
			}

			$val = $objResult.Properties[ 'info' ]
			If( $null -ne $val -and $val.Count -gt 0 )
			{
				$notes = $val.Item( 0 )
				$notes = $notes -replace "`r`n", "|"
			} 
			Else 
			{
				$notes = ''
			}

			If($IsStale) 
			{
				$TotalStaleObjects++
			}
			If($Enabled) 
			{
				$TotalEnabledObjects++
			}
			If($Enabled -eq $False) 
			{
				$TotalDisabledObjects++
			}
			If($IsStale -AND $Enabled) 
			{
				$TotalEnabledStaleObjects++
			}
			If($IsStale -AND $Enabled -eq $False) 
			{
				$TotalDisabledStaleObjects++
			}

			#V3.00 change to use PSCustomObject - MBS
			#$obj = New-Object -TypeName PSObject
			#$obj | Add-Member -MemberType NoteProperty -Name "Name" -value $Name
			#$obj | Add-Member -MemberType NoteProperty -Name "ParentOU" -value $ParentDN
			#$obj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -value $OperatingSystem
			#$obj | Add-Member -MemberType NoteProperty -Name "Version" -value $OperatingSystemVersion
			#$obj | Add-Member -MemberType NoteProperty -Name "ServicePack" -value $OperatingSystemServicePack
			#$obj | Add-Member -MemberType NoteProperty -Name "Description" -value $Description
			$obj = [PSCustomObject]@{
				"Name" = $Name
				"ParentOU" = $ParentDN
				"OperatingSystem" = $OperatingSystem
				"Version" = $OperatingSystemVersion
				"ServicePack" = $OperatingSystemServicePack
				"Description" = $Description
			}
			
			#V3.00 - simplify logic!
			If( $ServicePrincipalName -match 'MSClusterVirtualServer' )
			{
				$Category = 'CNO or VCO'
				$OperatingSystem = $OperatingSystem + ' - ' + $Category
			}
			Else
			{
				If( $OperatingSystem -match 'windows' )
				{
					If( $OperatingSystem -match 'server' )
					{
						$Category = 'Server'
					}
					Else
					{
						$Category = 'Workstation'
					}
				}
				Else
				{
					$Category = 'Other'
				}
			}
<#
			If(!($ServicePrincipalName -match 'MSClusterVirtualServer')) 
			{
				If($OperatingSystem -match 'windows' -AND $OperatingSystem -match 'server') 
				{
					$Category = "Server"
				}
				If($OperatingSystem -match 'windows' -AND !($OperatingSystem -match 'server')) 
				{
					$Category = "Workstation"
				}
				If(!($OperatingSystem -match 'windows')) 
				{
					$Category = "Other"
				}
			} 
			Else 
			{
				$Category = "CNO or VCO"
				$OperatingSystem = $OperatingSystem + " - " + $Category
			}
#>
			## FIXME - I don't believe this is possible - dead code - MBS
			If( $Category -eq '' ) 
			{
				$Category = 'Undefined'
			}

			$obj | Add-Member -MemberType NoteProperty -Name "Category" -value $Category
			$obj | Add-Member -MemberType NoteProperty -Name "PasswordLastSet" -value $PasswordLastSet
			$obj | Add-Member -MemberType NoteProperty -Name "LastLogon" -value $LastLogon
			$obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $Enabled
			$obj | Add-Member -MemberType NoteProperty -Name "IsStale" -value $IsStale
			$obj | Add-Member -MemberType NoteProperty -Name "WhenCreated" -value $WhenCreated
			$obj | Add-Member -MemberType NoteProperty -Name "Notes" -value $notes

			$null = $AllComputerObjects.Add( $obj )

			Switch($Category)
			{
				"Server"		{ $null = $WindowsServerObjects.Add(      $obj ); break }
				"Workstation"	{ $null = $WindowsWorkstationObjects.Add( $obj ); break }
				"Other"			{ $null = $NonWindowsComputerObjects.Add( $obj ); break }
				"CNO or VCO"	{ $null = $CNOandVCOObjects.Add(          $obj ); break }
				"Undefined"		{ $null = $NonWindowsComputerObjects.Add( $obj ); break }
			}
			$obj = New-Object -TypeName PSObject
			$obj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -value $OperatingSystem
			If($OperatingSystemIncludesServicePack -eq $False) 
			{
				$FullOperatingSystem = $OperatingSystem
			} 
			Else 
			{
				$FullOperatingSystem = $OperatingSystem + " " + $OperatingSystemServicePack
				$obj | Add-Member -MemberType NoteProperty -Name "ServicePack" -value $OperatingSystemServicePack
			}
			$obj | Add-Member -MemberType NoteProperty -Name "Category" -value $Category

			# Create a hashtable to capture a count of each Operating System
			If(!($ComputersHashTable.ContainsKey($FullOperatingSystem))) 
			{
				$TotalCount = 1
				$StaleCount = 0
				$EnabledStaleCount = 0
				$DisabledStaleCount = 0
				If($IsStale -eq $True) 
				{ 
					$StaleCount = 1 
				}
				If($Enabled -eq $True) 
				{
					$EnabledCount = 1
					$DisabledCount = 0
					If($IsStale -eq $True) 
					{ 
						$EnabledStaleCount = 1 
					}
				}
				If($Enabled -eq $False) 
				{
					$DisabledCount = 1
					$EnabledCount = 0
					If($IsStale -eq $True) 
					{ 
						$DisabledStaleCount = 1 
					}
				}
				$obj | Add-Member -MemberType NoteProperty -Name "Total" -value $TotalCount
				$obj | Add-Member -MemberType NoteProperty -Name "Stale" -value $StaleCount
				$obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $EnabledCount
				$obj | Add-Member -MemberType NoteProperty -Name "Enabled_Stale" -value $EnabledStaleCount
				$obj | Add-Member -MemberType NoteProperty -Name "Active" -value ($EnabledCount - $EnabledStaleCount)
				$obj | Add-Member -MemberType NoteProperty -Name "Disabled" -value $DisabledCount
				$obj | Add-Member -MemberType NoteProperty -Name "Disabled_Stale" -value $DisabledStaleCount
				$ComputersHashTable = $ComputersHashTable + @{$FullOperatingSystem = $obj}  ## FIXME - whut??? - MBS
			} 
			Else 
			{
				$value = $ComputersHashTable.Get_Item($FullOperatingSystem)
				$TotalCount = $value.Total + 1
				$StaleCount = $value.Stale
				$EnabledStaleCount = $value.Enabled_Stale
				$DisabledStaleCount = $value.Disabled_Stale
				If($IsStale -eq $True) 
				{ 
					$StaleCount = $value.Stale + 1 
				}
				If($Enabled -eq $True) 
				{
					$EnabledCount = $value.Enabled + 1
					$DisabledCount = $value.Disabled
					If($IsStale -eq $True) 
					{ 
						$EnabledStaleCount = $value.Enabled_Stale + 1 
					}
				}
				If($Enabled -eq $False) 
				{ 
					$DisabledCount = $value.Disabled + 1
					$EnabledCount = $value.Enabled
					If($IsStale -eq $True) 
					{ 
						$DisabledStaleCount = $value.Disabled_Stale + 1 
					}
				}
				$obj | Add-Member -MemberType NoteProperty -Name "Total" -value $TotalCount
				$obj | Add-Member -MemberType NoteProperty -Name "Stale" -value $StaleCount
				$obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $EnabledCount
				$obj | Add-Member -MemberType NoteProperty -Name "Enabled_Stale" -value $EnabledStaleCount
				$obj | Add-Member -MemberType NoteProperty -Name "Active" -value ($EnabledCount - $EnabledStaleCount)
				$obj | Add-Member -MemberType NoteProperty -Name "Disabled" -value $DisabledCount
				$obj | Add-Member -MemberType NoteProperty -Name "Disabled_Stale" -value $DisabledStaleCount
				$ComputersHashTable.Set_Item($FullOperatingSystem,$obj)   ### FIXME - whut??? - MBS
			} # end if
			$TotalComputersProcessed ++
		}

		# Dispose of the search and results properly to avoid a memory leak
		$colResults.Dispose()

		$Output = $ComputersHashTable.values | ForEach-Object {$_ } | ForEach-Object {$_ } | Sort-Object OperatingSystem -descending
		
		If($MSWORD -or $PDF)
		{
			WriteWordLine 3 0 "Computer Operating Systems"
		}
		If($Text)
		{
			Line 0 "Computer Operating Systems"
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 "Computer Operating Systems"
		}

		ForEach($Item in $Output)
		{
			If($MSWORD -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Operating System"; Value = $Item.OperatingSystem; }
				$ScriptInformation += @{ Data = "Service Pack"; Value = $Item.ServicePack; }
				$ScriptInformation += @{ Data = "Category"; Value = $Item.Category; }
				$ScriptInformation += @{ Data = "Total"; Value = $Item.Total; }
				$ScriptInformation += @{ Data = "Stale"; Value = $Item.Stale; }
				$ScriptInformation += @{ Data = "Enabled"; Value = $Item.Enabled; }
				$ScriptInformation += @{ Data = "Enabled/Stale"; Value = $Item.Enabled_Stale; }
				$ScriptInformation += @{ Data = "Active"; Value = $Item.Active; }
				$ScriptInformation += @{ Data = "Disabled"; Value = $Item.Disabled; }
				$ScriptInformation += @{ Data = "Disabled/Stale"; Value = $Item.Disabled_Stale; }
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				## IB - set column widths without recursion
				$Table.Columns.Item(1).Width = 100;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Operating System`t: " $Item.OperatingSystem
				Line 1 "Service Pack`t`t: " $Item.ServicePack
				Line 1 "Category`t`t: " $Item.Category
				Line 1 "Total`t`t`t: " $Item.Total
				Line 1 "Stale`t`t`t: " $Item.Stale
				Line 1 "Enabled`t`t`t: " $Item.Enabled
				Line 1 "Enabled/Stale`t`t: " $Item.Enabled_Stale
				Line 1 "Active`t`t`t: " $Item.Active
				Line 1 "Disabled`t`t: " $Item.Disabled
				Line 1 "Disabled/Stale`t`t: " $Item.Disabled_Stale
				Line 0 ""
			}
			If($HTML)
			{
				#V3.00 pre-allocate rowdata
				## $rowdata = @()
				$rowdata = New-Object System.Array[] 9 

				$rowdata[ 0 ] = @(
					'Service Pack',    $htmlsb,
					$Item.ServicePack, $htmlwhite
				)
				$rowdata[ 1 ] = @(
					'Category',     $htmlsb,
					$Item.Category, $htmlwhite
				)
				$rowdata[ 2 ] = @(
					'Total',     $htmlsb,
					$Item.Total.ToString(), $htmlwhite
				)
				$rowdata[ 3 ] = @(
					'Stale',     $htmlsb,
					$Item.Stale.ToString(), $htmlwhite
				)
				$rowdata[ 4 ] = @(
					'Enabled',     $htmlsb,
					$Item.Enabled.ToString(), $htmlwhite
				)
				$rowdata[ 5 ] = @(
					'Enabled/Stale',     $htmlsb,
					$Item.Enabled_Stale.ToString(), $htmlwhite
				)
				$rowdata[ 6 ] = @(
					'Active',     $htmlsb,
					$Item.Active.ToString(), $htmlwhite
				)
				$rowdata[ 7 ] = @(
					'Disabled',     $htmlsb,
					$Item.Disabled.ToString(), $htmlwhite
				)
				$rowdata[ 8 ] = @(
					'Disabled/Stale',     $htmlsb,
					$Item.Disabled_Stale.ToString(), $htmlwhite
				)
				
				$columnWidths  = @( '150px', '250px' )
				$columnHeaders = @(
					'Operating System', $htmlsb,
					$Item.OperatingSystem, $htmlwhite
				)

				FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '400'
				WriteHTMLLine 0 0 ''

				$rowdata = $null
			}
		}

		$percent = "{0:P}" -f ($TotalStaleObjects/$ComputerCount)
		
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "A breakdown of the $ComputerCount Computer Objects in the $domain Domain"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Total Computer Objects"; Value = $ComputerCount; }
			$ScriptInformation += @{ Data = "Total Stale Computer Objects (count)"; Value = $TotalStaleObjects; }
			$ScriptInformation += @{ Data = "Total Stale Computer Objects (percent)"; Value = $percent; }
			$ScriptInformation += @{ Data = "Total Enabled Computer Objects"; Value = $TotalEnabledObjects; }
			$ScriptInformation += @{ Data = "Total Enabled Stale Computer Objects"; Value = $TotalEnabledStaleObjects; }
			$ScriptInformation += @{ Data = "Total Active Computer Objects"; Value = $($TotalEnabledObjects - $TotalEnabledStaleObjects); }
			$ScriptInformation += @{ Data = "Total Disabled Computer Objects"; Value = $TotalDisabledObjects; }
			$ScriptInformation += @{ Data = "Total Disabled Stale Computer Objects"; Value = $TotalDisabledStaleObjects; }
			
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 "A breakdown of the $ComputerCount Computer Objects in the $domain Domain"
			
			Line 1 "Total Computer Objects`t`t`t: " $ComputerCount
			Line 1 "Total Stale Computer Objects (count)`t: " $TotalStaleObjects
			Line 1 "Total Stale Computer Objects (percent)`t: " $percent
			Line 1 "Total Enabled Computer Objects`t`t: " $TotalEnabledObjects
			Line 1 "Total Enabled Stale Computer Objects`t: " $TotalEnabledStaleObjects
			Line 1 "Total Active Computer Objects`t`t: " $($TotalEnabledObjects - $TotalEnabledStaleObjects)
			Line 1 "Total Disabled Computer Objects`t`t: " $TotalDisabledObjects
			Line 1 "Total Disabled Stale Computer Objects`t: " $TotalDisabledStaleObjects
			Line 0 ""
		}
		If($HTML)
		{
			#V3.00 - pre-allocate rowdata
			## $rowdata = @()
			$rowdata = New-Object System.Array[] 7

			$rowdata[ 0 ] = @(
				'Total Stale Computer Objects (count)', $htmlsb,
				$TotalStaleObjects.ToString(), ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
			)
			$rowdata[ 1 ] = @(
				'Total Stale Computer Objects (percent)',$htmlsb,
				$percent, ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
			)
			$rowdata[ 2 ] = @(
				'Total Enabled Computer Objects', $htmlsb,
				$TotalEnabledObjects.ToString(), ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
			)
			$rowdata[ 3 ] = @(
				'Total Enabled Stale Computer Objects', $htmlsb,
				$TotalEnabledStaleObjects.ToString(), ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
			)
			$rowdata[ 4 ] = @(
				'Total Active Computer Objects', $htmlsb,
				$( $TotalEnabledObjects - $TotalEnabledStaleObjects ), ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
			)
			$rowdata[ 5 ] = @(
				'Total Disabled Computer Objects', $htmlsb,
				$TotalDisabledObjects.ToString(), ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
			)
			$rowdata[ 6 ] = @(
				'Total Disabled Stale Computer Objects', $htmlsb,
				$TotalDisabledStaleObjects.ToString(), ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
			)

			$msg           = "A breakdown of the $ComputerCount Computer Objects in the $domain Domain"
			$columnWidths  = @( '400px', '50px' )
			$columnHeaders = @(
				'Total Computer Objects',  $htmlsb,
				$ComputerCount.ToString(), ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
			)

			FormatHTMLTable -tableHeader $msg `
				-rowArray $rowdata `
				-columnArray $columnHeaders `
				-fixedWidth $columnWidths `
				-tablewidth '450'
			WriteHTMLLine 0 0 ''

			$rowData = $null
		}

		#Notes:
		# - Computer objects are filtered into 4 categories:
		#	 - Windows Servers
		#	 - Windows Workstations
		#	 - Other non-Windows (Linux, Mac, etc)
		#	 - CNO or VCO (Windows Cluster Name Objects and Virtual Computer Objects)
		# - A Stale object is derived from 2 values ANDed together:
		#     PasswordLastChanged  > $MaxPasswordLastChanged days ago
		#     AND
		#     LastLogonDate > $MaxLastLogonDate days ago
		# - If it's never logged on before and was created more than $MaxLastLogonDate days ago, set the
		#   HasNotRecentlyLoggedOn variable to True. This will also be used to help determine if
		#   it's a stale account. An example of this would be if you prestaged the account but
		#   never ended up using it.
		# - The Active objects column is calculated by subtracting the Enabled_Stale value from
		#   the Enabled value. This gives us an accurate number of active objects against each
		#   Operating System.
		# - To help provide a high level overview of the computer object landscape we calculate
		#   the number of stale objects of enabled and disabled objects separately.
		#   Disabled objects are often ignored, but it's pointless leaving old disabled computer
		#   objects in the domain.
		# - For viewing purposes we sort the output by Operating System and not count.
		# - You may notice a question mark (?) in some of the OperatingSystem strings. This is a
		#   representation of each Double-Byte character that was unable to be translated. Refer
		#   to Microsoft KB829856 for an explanation.
		# - Be aware that a cluster updates the lastLogonTimeStamp of the CNO/VNO when it brings
		#   a clustered network name resource online. So it could be running for months without
		#   an update to the lastLogonTimeStamp attribute.
		# - When a VMware ESX host has been added to Active Directory its associated computer
		#   object will appear as an Operating System of unknown, version: unknown, and service
		#   pack: Likewise Identity 5.3.0. The lsassd.conf manages the machine password expiration
		#   lifespan, which by default is set to 30 days.
		# - When a Riverbed SteelHead has been added to Active Directory the password set on the
		#   Computer Object will never change by default. This must be enabled on the command
		#   line using the 'domain settings pwd-refresh-int <number of days>' command. The
		#   lastLogon timestamp is only updated when the SteelHead appliance is restarted.
	}
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): AddDateTime     : $AddDateTime"
	Write-Verbose "$(Get-Date -Format G): Company Name    : $Script:CoName"
	Write-Verbose "$(Get-Date -Format G): ComputerName    : $ComputerName"	#3.06 always output
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Company Address : $CompanyAddress"
		Write-Verbose "$(Get-Date -Format G): Company Email   : $CompanyEmail"
		Write-Verbose "$(Get-Date -Format G): Company Fax     : $CompanyFax"
		Write-Verbose "$(Get-Date -Format G): Company Phone   : $CompanyPhone"
		Write-Verbose "$(Get-Date -Format G): Cover Page      : $CoverPage"
	}
	Write-Verbose "$(Get-Date -Format G): DCDNSInfo       : $DCDNSInfo"
	Write-Verbose "$(Get-Date -Format G): Dev             : $Dev"
	If($Dev)
	{
		Write-Verbose "$(Get-Date -Format G): DevErrorFile    : $Script:DevErrorFile"
	}
	Write-Verbose "$(Get-Date -Format G): Domain Name     : $ADDomain"
	Write-Verbose "$(Get-Date -Format G): Elevated        : $Script:Elevated"
	If($MSWord)
	{
		Write-Verbose "$(Get-Date -Format G): Word FileName   : $($Script:WordFileName)"
	}
	If($HTML)
	{
		Write-Verbose "$(Get-Date -Format G): HTML FileName   : $($Script:HTMLFileName)"
	} 
	If($PDF)
	{
		Write-Verbose "$(Get-Date -Format G): PDF FileName    : $($Script:PDFFileName)"
	}
	If($Text)
	{
		Write-Verbose "$(Get-Date -Format G): Text FileName   : $($Script:TextFileName)"
	}
	Write-Verbose "$(Get-Date -Format G): Folder          : $Folder"
	Write-Verbose "$(Get-Date -Format G): Forest Name     : $ADForest"
	Write-Verbose "$(Get-Date -Format G): From            : $From"
	Write-Verbose "$(Get-Date -Format G): GPOInheritance  : $GPOInheritance"
	Write-Verbose "$(Get-Date -Format G): HW Inventory    : $Hardware"
	Write-Verbose "$(Get-Date -Format G): IncludeUserInfo : $IncludeUserInfo"
	Write-Verbose "$(Get-Date -Format G): Log             : $($Log)"
	Write-Verbose "$(Get-Date -Format G): MaxDetail       : $MaxDetails"
	Write-Verbose "$(Get-Date -Format G): Report Footer   : $ReportFooter"
	Write-Verbose "$(Get-Date -Format G): Save As HTML    : $HTML"
	Write-Verbose "$(Get-Date -Format G): Save As PDF     : $PDF"
	Write-Verbose "$(Get-Date -Format G): Save As TEXT    : $TEXT"
	Write-Verbose "$(Get-Date -Format G): Save As WORD    : $MSWORD"
	Write-Verbose "$(Get-Date -Format G): ScriptInfo      : $ScriptInfo"
	Write-Verbose "$(Get-Date -Format G): Section         : $Section"
	Write-Verbose "$(Get-Date -Format G): Services        : $Services"
	Write-Verbose "$(Get-Date -Format G): Smtp Port       : $SmtpPort"
	Write-Verbose "$(Get-Date -Format G): Smtp Server     : $SmtpServer"
	Write-Verbose "$(Get-Date -Format G): Title           : $Script:Title"
	Write-Verbose "$(Get-Date -Format G): To              : $To"
	Write-Verbose "$(Get-Date -Format G): Use SSL         : $UseSSL"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): User Name       : $UserName"
	}
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): OS Detected     : $Script:RunningOS"
	Write-Verbose "$(Get-Date -Format G): PoSH version    : $($Host.Version)"
	Write-Verbose "$(Get-Date -Format G): PSCulture       : $PSCulture"
	Write-Verbose "$(Get-Date -Format G): PSUICulture     : $PSUICulture"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Word language   : $Script:WordLanguageValue"
		Write-Verbose "$(Get-Date -Format G): Word version    : $Script:WordProduct"
	}
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): Script start    : $Script:StartTime"
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date -Format G): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:WordFileName, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:WordFileName, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date -Format G): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	Write-Verbose "$(Get-Date -Format G): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword Process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword running in our session
	$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}) | Select-Object -Property Id 
	If( $wordprocess -and $wordprocess.Id -gt 0)
	{
		Write-Verbose "$(Get-Date -Format G): WinWord Process is still running. Attempting to stop WinWord Process # $($wordprocess.Id)"
		Stop-Process $wordprocess.Id -EA 0
	}
}

Function SetupText
{
	Write-Verbose "$(Get-Date -Format G): Setting up Text"

	[System.Text.StringBuilder] $global:Output = New-Object System.Text.StringBuilder( 16384 )

	If(!$AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName).txt"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}
}

Function SaveandCloseTextDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving Text file"
	Line 0 ""
	Line 0 "Report Complete"
	Write-Output $global:Output.ToString() | Out-File $Script:TextFileName 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving HTML file"
	WriteHTMLLine 0 0 ""
	WriteHTMLLine 0 0 "Report Complete"
	Out-File -FilePath $Script:HTMLFileName -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFilenames
{
	Param([string]$OutputFileName)
	
	If($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		SetupWord
	}
	If($Text)
	{
		SetupText
	}
	If($HTML)
	{
		SetupHTML
	}
	ShowScriptOptions
}
#endregion

#Script begins

#region script setup Function
Function ProcessScriptSetup
{
	#If hardware inventory or services are requested, make sure user is running the script with Domain Admin rights
	Write-Verbose "$(Get-Date -Format G): Testing to see if $env:username has Domain Admin rights"
	
	$AmIReallyDA = UserIsaDomainAdmin
	If($AmIReallyDA -eq $True)
	{
		#user has Domain Admin rights
		If($ADDomain -ne "")
		{
			Write-Verbose "$(Get-Date -Format G): `t$env:username has Domain Admin rights in the $ADDomain Domain"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): `t$env:username has Domain Admin rights in the $ADForest Forest"
		}
		$Script:DARights = $True
	}
	Else
	{
		#user does not have Domain Admin rights
		If($ADDomain -ne "")
		{
			Write-Verbose "$(Get-Date -Format G): `t$env:username does not have Domain Admin rights in the $ADDomain Domain"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): `t$env:username does not have Domain Admin rights in the $ADForest Forest"
		}
	}
	
	$Script:Elevated = ElevatedSession
	
	If($Hardware -or $Services -or $DCDNSInfo)
	{
		If($Hardware)
		{
			Write-Verbose "$(Get-Date -Format G): Hardware inventory requested"
		}

		If($Services)
		{
			Write-Verbose "$(Get-Date -Format G): Services requested"
		}
		
		If($DCDNSInfo)
		{
			Write-Verbose "$(Get-Date -Format G): Domain Controller DNS configuration information requested"
		}

#The following write-host statements should NOT be indented in the code or it messes up how they look in the console when the script runs
		If($Script:DARights -eq $False)
		{
			#user does not have Domain Admin rights
			If($Hardware)
			{
				#don't abort script, set $hardware to false
				Write-Host "
Hardware inventory was requested but $($env:username) does not have Domain Admin rights.
Hardware inventory option will be turned off." -Foreground White
				$Script:Hardware = $False
			}

			If($Services)
			{
				#don't abort script, set $services to false
				Write-Host "
Services were requested but $($env:username) does not have Domain Admin rights.
Services option will be turned off." -Foreground White
				$Script:Services = $False
			}
		}
		
		If($Script:Elevated -eq $False)
		{
			#user is not running the script from an elevated PoSH session
			If($Hardware)
			{
				#don't abort script, set $hardware to false
				Write-Host "
Hardware inventory was requested but this is not an elevated PowerShell session.
Hardware inventory option will be turned off." -Foreground White
				$Script:Hardware = $False
			}

			If($Services)
			{
				#don't abort script, set $services to false
				Write-Host "
Services were requested but this is not an elevated PowerShell session.
Services option will be turned off." -Foreground White
				$Script:Services = $False
			}
		}

		If(!$Script:DARights -and !$Script:Elevated)
		{
			Write-Host "
To obtain Time Server and AD file location data,
please run the script from an elevated PowerShell session using an account with Domain Admin rights.
" -Foreground White
		}
	}

	#if computer name is localhost, get actual server name
	If($ComputerName -eq "localhost")
	{
		$Script:ComputerName = $env:ComputerName
		Write-Verbose "$(Get-Date -Format G): Server name has been changed from localhost to $($ComputerName)"
	}
	
	#see if default value of $Env:USERDNSDOMAIN was used
	If($ComputerName -eq $Env:USERDNSDOMAIN)
	{
		#change $ComputerName to a found global catalog server
		$Results = (Get-ADDomainController -DomainName $ADForest -Discover -Service GlobalCatalog -EA 0).Name
		
		If($? -and $Null -ne $Results)
		{
			$Script:ComputerName = $Results
			Write-Verbose "$(Get-Date -Format G): Server name has been changed from $Env:USERDNSDOMAIN to $ComputerName"
		}
		ElseIf(!$?) #changed for 2.16
		{
			#may be in a child domain where -Service GlobalCatalog doesn't work. Try PrimaryDC
			$Results = (Get-ADDomainController -DomainName $ADForest -Discover -Service PrimaryDC -EA 0).Name

			If($? -and $Null -ne $Results)
			{
				$Script:ComputerName = $Results
				Write-Verbose "$(Get-Date -Format G): Server name has been changed from $Env:USERDNSDOMAIN to $ComputerName"
			}
		}
	}

	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $ComputerName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Null -ne $Result)
		{
			$Script:ComputerName = $Result.HostName
			Write-Verbose "$(Get-Date -Format G): Server name has been changed from $ip to $ComputerName"
		}
		Else
		{
			Write-Warning "Unable to resolve $ComputerName to a hostname"
		}
	}
	Else
	{
		#server is online but for some reason $ComputerName cannot be converted to a System.Net.IpAddress
	}

	If(![String]::IsNullOrEmpty($ComputerName)) 
	{
		#get server name
		#first test to make sure the server is reachable
		Write-Verbose "$(Get-Date -Format G): Testing to see if $ComputerName is online and reachable"
		#If(Test-Connection -ComputerName $ComputerName -quiet -EA 0)
		If(Test-NetConnection -ComputerName $ComputerName -Port 88 -InformationLevel Quiet -EA 0) #port 88 is the KDC and is unique to DCs (thanks to Matthew Woolnough)
		{
			Write-Verbose "$(Get-Date -Format G): `tServer $ComputerName is online."
			Write-Verbose "$(Get-Date -Format G): `tTest #1 to see if $ComputerName is a Domain Controller."
			#the server may be online but is it really a domain controller?

			#is the ComputerName in the specified forest or domain
			$Results = Get-ADDomainController -Server $ComputerName -EA 0
			
			If(!$? -or $Null -eq $Results)
			{
				#try using the Forest name
				Write-Verbose "$(Get-Date -Format G): `tTest #2 to see if $ComputerName is a Domain Controller."
				
				$Results = Get-ADDomainController -Server $ComputerName -Domain $ADForest -EA 0
				
				If(!$?)
				{
					$ErrorActionPreference = $SaveEAPreference
					Write-Error "
					`n`n
					`t`t
					$ComputerName is not a domain controller for $ADForest.
					`n`n
					`t`t
					Script cannot Continue.
					`n`n
					"
					AbortScript
				}
				Else
				{
					Write-Verbose "$(Get-Date -Format G): `tTest #2 succeeded. $ComputerName is a Domain Controller."
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): `tTest #1 succeeded. $ComputerName is a Domain Controller."
			}
			
			$Results = $Null
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Computer $ComputerName is offline"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "
			`n`n
			`t`t
			Computer $ComputerName is offline.
			`n`n
			`t`t
			Script cannot Continue.
			`n`n
			"
			AbortScript
		}
	}

	If($ADForest -ne $ADDomain)
	{
		#get forest information so output filename can be generated
		Write-Verbose "$(Get-Date -Format G): Testing to see if $($ADForest) is a valid forest name"
		If([String]::IsNullOrEmpty($ComputerName))
		{
			$Script:Forest = Get-ADForest -Identity $ADForest -EA 0
			
			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a forest identified by: $ADForest.
				`n`n
				`t`t
				Script cannot Continue.
				`n`n
				"
				AbortScript
			}
		}
		Else
		{
			$Script:Forest = Get-ADForest -Identity $ADForest -Server $ComputerName -EA 0

			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a forest with the name of $ADForest.
				`n`n
				`t`t
				Script cannot Continue.
				`n`n
				`t`t
				Is $ComputerName running Active Directory Web Services?
				`n`n
				"
				AbortScript
			}
		}
		
		#3.04 Make sure $Forest.Name matches the name passsed to -ADForest
		If($ADForest -eq $Script:Forest.Name)
		{
			Write-Verbose "$(Get-Date -Format G): `t$ADForest is a valid forest name"
			[string]$Script:Title = "AD Inventory Report for the $ADForest Forest"
			$Script:Domains       = $Script:Forest.Domains | Sort-Object 
			$Script:ConfigNC      = (Get-ADRootDSE -Server $ADForest -EA 0).ConfigurationNamingContext
		}
		Else
		{
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "
		`n`n
		$ADForest is not the name of the Forest's Root Domain of $($Script:Forest.Name).
		`n`n
		Run the script using -ADDomain instead of -ADForest
		`n`n
		Script cannot Continue.
		`n`n
			"
			AbortScript
		}
	}
	
	If($ADDomain -ne "")
	{
		Write-Verbose "$(Get-Date -Format G): Testing to see if $($ADDomain) is a valid domain name"
		If([String]::IsNullOrEmpty($ComputerName))
		{
			$results = Get-ADDomain -Identity $ADDomain -EA 0
			
			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a domain identified by: $ADDomain.
				`n`n
				`t`t
				Script cannot Continue.
				`n`n
				"
				AbortScript
			}
		}
		Else
		{
			$results = Get-ADDomain -Identity $ADDomain -Server $ComputerName -EA 0

			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a domain with the name of $ADDomain.
				`n`n
				`t`t
				Script cannot Continue.
				`n`n
				`t`t
				Is $ComputerName running Active Directory Web Services?
				`n`n
				"
				AbortScript
			}
		}
		Write-Verbose "$(Get-Date -Format G): `t$ADDomain is a valid domain name"
		$Script:Domains       = $results.DNSRoot
		$Script:DomainDNSRoot = $results.DNSRoot
		[string]$Script:Title = "AD Inventory Report for the $Script:Domains Domain"
		
		$tmp = $results.Forest
		#get forest info 
		Write-Verbose "$(Get-Date -Format G): Retrieving forest information"
		If([String]::IsNullOrEmpty($ComputerName))
		{
			$Script:Forest = Get-ADForest -Identity $tmp -EA 0
			
			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a forest identified by: $tmp.
				`n`n
				`t`t
				Script cannot Continue.
				`n`n
				"
				AbortScript
			}
		}
		Else
		{
			$Script:Forest = Get-ADForest -Identity $tmp -Server $ComputerName -EA 0

			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a forest with the name of $tmp.
				`n`n
				`t`t
				Script cannot Continue.
				`n`n
				`t`t
				Is $ComputerName running Active Directory Web Services?
				`n`n
				"
				AbortScript
			}
		}
		Write-Verbose "$(Get-Date -Format G): `tFound forest information for $tmp"
		$Script:ConfigNC = (Get-ADRootDSE -Server $tmp -EA 0).ConfigurationNamingContext
	}
	
	#store root domain so it only has to be accessed once
	[string]$Script:ForestRootDomain = $Script:Forest.RootDomain
	[string]$Script:ForestName       = $Script:Forest.Name
	#set naming context
}
#endregion

######################START OF BUILDING REPORT

#region Forest information
Function ProcessForestInformation
{
	Write-Verbose "$(Get-Date -Format G): Writing forest data"
	ProcessScriptStart
	
	If(![String]::IsNullOrEmpty($Script:CoName))
	{
		$txt = "Forest Information for $Script:CoName"
	}
	Else
	{
		$txt = "Forest Information"
	}

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Forest Information"
	}
	If($Text)
	{
		Line 0 "///  $txt  \\\"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;$txt&nbsp;&nbsp;\\\"
	}

	Switch ($Script:Forest.ForestMode)
	{
		"0"	{$ForestMode = "Windows 2000"; Break}
		"1" {$ForestMode = "Windows Server 2003 interim"; Break}
		"2" {$ForestMode = "Windows Server 2003"; Break}
		"3" {$ForestMode = "Windows Server 2008"; Break}
		"4" {$ForestMode = "Windows Server 2008 R2"; Break}
		"5" {$ForestMode = "Windows Server 2012"; Break}
		"6" {$ForestMode = "Windows Server 2012 R2"; Break}
		"7" {$ForestMode = "Windows Server 2016"; Break}	#added V2.20
		"Windows2000Forest"        {$ForestMode = "Windows 2000"; Break}
		"Windows2003InterimForest" {$ForestMode = "Windows Server 2003 interim"; Break}
		"Windows2003Forest"        {$ForestMode = "Windows Server 2003"; Break}
		"Windows2008Forest"        {$ForestMode = "Windows Server 2008"; Break}
		"Windows2008R2Forest"      {$ForestMode = "Windows Server 2008 R2"; Break}
		"Windows2012Forest"        {$ForestMode = "Windows Server 2012"; Break}
		"Windows2012R2Forest"      {$ForestMode = "Windows Server 2012 R2"; Break}
		"WindowsThresholdForest"   {$ForestMode = "Windows Server 2016 TP4"; Break}
		"Windows2016Forest"		   {$ForestMode = "Windows Server 2016"; Break}
		"UnknownForest"            {$ForestMode = "Unknown Forest Mode"; Break}
		Default                    {$ForestMode = "Unable to determine Forest Mode: $($Script:Forest.ForestMode)"; Break}
	}

	$AppPartitions         = $Script:Forest.ApplicationPartitions | Sort-Object 
	$CrossForestReferences = $Script:Forest.CrossForestReferences | Sort-Object 
	$SPNSuffixes           = $Script:Forest.SPNSuffixes | Sort-Object 
	$UPNSuffixes           = $Script:Forest.UPNSuffixes | Sort-Object 
	$Sites                 = $Script:Forest.Sites | Sort-Object 
	
	#added 9-oct-2016
	#https://adsecurity.org/?p=81
	$DirectoryServicesConfigPartition = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$Script:ConfigNC" -Partition $Script:ConfigNC -Properties *
	
	$TombstoneLifetime = $DirectoryServicesConfigPartition.tombstoneLifetime
	
	If($Null -eq $TombstoneLifetime -or $TombstoneLifetime -eq 0)
	{
		$TombstoneLifetime = 60
	}

	#2.16
	#move this duplicated block of code outside the output format test
	If($ADDomain -ne "")
	{
		#2.16 don't mess with the $Script:Domains variable
		#redo list of domains so forest root domain is listed first
		[array]$tmpDomains = "$Script:ForestRootDomain"
		[array]$tmpDomains2 = "$($Script:ForestRootDomain)"
		ForEach($Domain in $Forest.Domains)
		{
			If($Domain -ne $Script:ForestRootDomain)
			{
				$tmpDomains += "$($Domain.ToString())"
				$tmpDomains2 += "$($Domain.ToString())"
			}
		}
	}
	Else
	{
		#redo list of domains so forest root domain is listed first
		[array]$tmpDomains = "$Script:ForestRootDomain"
		[array]$tmpDomains2 = "$($Script:ForestRootDomain)"
		ForEach($Domain in $Script:Domains)
		{
			If($Domain -ne $Script:ForestRootDomain)
			{
				$tmpDomains += "$($Domain.ToString())"
				$tmpDomains2 += "$($Domain.ToString())"
			}
		}
		
		$Script:Domains = $tmpDomains
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Forest mode"; Value = $ForestMode; }
		$ScriptInformation += @{ Data = "Forest name"; Value = $Script:Forest.Name; }
		$tmp = ""
		If($Null -eq $AppPartitions)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "Application partitions"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($AppPartition in $AppPartitions)
			{
				$cnt++
				$tmp = "$($AppPartition.ToString())"
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "Application partitions"; Value = $tmp; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$tmp = ""
		If($Null -eq $CrossForestReferences)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "Cross forest references"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($CrossForestReference in $CrossForestReferences)
			{
				$cnt++
				$tmp = "$($CrossForestReference.ToString())"
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "Cross forest references"; Value = $tmp; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$ScriptInformation += @{ Data = "Domain naming master"; Value = $Script:Forest.DomainNamingMaster; }
		$tmp = ""
		If($Null -eq $Script:Domains)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "Domains in forest"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($Domain in $tmpDomains2)
			{
				$cnt++
				$tmp = "$($Domain.ToString())"
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "Domains in forest"; Value = $tmp; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$ScriptInformation += @{ Data = "Partitions container"; Value = $Script:Forest.PartitionsContainer; }
		$ScriptInformation += @{ Data = "Root domain"; Value = $Script:ForestRootDomain; }
		$ScriptInformation += @{ Data = "Schema master"; Value = $Script:Forest.SchemaMaster; }
		$tmp = ""
		If($Null -eq $Sites)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "Sites"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($Site in $Sites)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "Sites"; Value = $Site.ToString(); }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $Site.ToString(); }
				}
			}
		}
		$tmp = ""
		If($Null -eq $SPNSuffixes)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "SPN suffixes"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($SPNSuffix in $SPNSuffixes)
			{
				$cnt++
				$tmp = "$($SPNSuffix.ToString())"
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "SPN suffixes"; Value = $tmp; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$ScriptInformation += @{ Data = "Tombstone lifetime"; Value = "$($TombstoneLifetime) days"; }
		$tmp = ""
		If($Null -eq $UPNSuffixes)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "UPN suffixes"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($UPNSuffix in $UPNSuffixes)
			{
				$cnt++
				$tmp = "$($UPNSuffix.ToString())"
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "UPN suffixes"; Value = $tmp; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$tmp = $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 300;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0  "Forest mode`t`t: " $ForestMode
		Line 0  "Forest name`t`t: " $Script:Forest.Name
		If($Null -eq $AppPartitions)
		{
			Line 0 "Application partitions`t: <None>"
		}
		Else
		{
			Line 0 "Application partitions`t: " -NoNewLine
			$cnt = 0
			ForEach($AppPartition in $AppPartitions)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 "$($AppPartition.ToString())"
				}
				Else
				{
					Line 3 "  $($AppPartition.ToString())"
				}
			}
		}
		If($Null -eq $CrossForestReferences)
		{
			Line 0 "Cross forest references`t: <None>"
		}
		Else
		{
			Line 0 "Cross forest references`t: " -NoNewLine
			$cnt = 0
			ForEach($CrossForestReference in $CrossForestReferences)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 "$($CrossForestReference.ToString())"
				}
				Else
				{
					Line 3 "  $($CrossForestReference.ToString())"
				}
			}
		}
		Line 0  "Domain naming master`t: " $Script:Forest.DomainNamingMaster
		If($Null -eq $Script:Domains)
		{
			Line 0 "Domains in forest`t: <None>"
		}
		Else
		{
			Line 0 "Domains in forest`t: " -NoNewLine
			$cnt = 0
			
			ForEach($Domain in $tmpDomains2)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 $Domain
				}
				Else
				{
					Line 3 "  " $Domain
				}
			}
		}
		Line 0  "Partitions container`t: " $Script:Forest.PartitionsContainer
		Line 0  "Root domain`t`t: " $Script:ForestRootDomain
		Line 0  "Schema master`t`t: " $Script:Forest.SchemaMaster
		If($Null -eq $Sites)
		{
			Line 0 "Sites`t`t`t: <None>"
		}
		Else
		{
			Line 0 "Sites`t`t`t: " -NoNewLine
			$cnt = 0
			ForEach($Site in $Sites)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 $Site.ToString()
				}
				Else
				{
					Line 3 "  $($Site.ToString())"
				}
			}
		}
		If($Null -eq $SPNSuffixes)
		{
			Line 0 "SPN suffixes`t`t: <None>"
		}
		Else
		{
			Line 0 "SPN suffixes`t`t: " -NoNewLine
			$cnt = 0
			ForEach($SPNSuffix in $SPNSuffixes)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 "$($SPNSuffix.ToString())"
				}
				Else
				{
					Line 3 "  $($SPNSuffix.ToString())"
				}
			}
		}
		Line 0 "Tombstone lifetime`t: " "$($TombstoneLifetime) days"
		If($Null -eq $UPNSuffixes)
		{
			Line 0 "UPN Suffixes`t`t: <None>"
		}
		Else
		{
			Line 0 "UPN Suffixes`t`t: " -NoNewLine
			$cnt = 0
			ForEach($UPNSuffix in $UPNSuffixes)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 "$($UPNSuffix.ToString())"
				}
				Else
				{
					Line 3 "  $($UPNSuffix.ToString())"
				}
			}
		}
		Line 0 ""
	}
	If($HTML)
	{
		#V3.00 - pre-allocate rowdata
		## $rowdata = @()
		#V3.00 - as you can see, it was a bit challenging to pre-allocate rowData
		#V3.00 - add $ForestMode is first line/columnHeader to match PDF/Word/Text
		## rowsRequired = 1 + ## forest name
		## MAX( 1, $AppPartitions.Count ) +
		## MAX( 1, $CrossForestReferences.Count ) +
		## 1 + ## domain naming master
		## MAX( 1, $tmpDomains2.Count ) + ## domains in forest
		## 3 + ## partitions container, root domain, schema master
		## MAX( 1, $Sites.Count ) +
		## MAX( 1, $SPNSuffixes.Count ) +
		## 1 + ## tombstone lifetime
		## MAX( 1, $UPNSuffixes.Count )
		Function valMAX
		{
			Param
			(
				[Object] $object
			)

			If( $object -and $object -is [Array] )
			{
				$object.Count
			}
			Else
			{
				1
			}
		}

		$rowsRequired = 1 +
			$( valMax $AppPartitions         ) +
			$( valMax $CrossForestReferences ) +
			1 +
			$( valMax $tmpDomains2           ) + #V3.04 changed from $Script:domains to $tmpDomains2. We want all domains in the forest
			3 +
			$( valMax $Sites                 ) +
			$( valMax $SPNSuffixes           ) +
			1 +
			$( valMax $UPNSuffixes           ) +
			0
		$rowdata  = New-Object System.Array[] $rowsRequired
		$rowIndex = 0

		$rowdata[ $rowIndex++ ] = @(
			'Forest name',       $htmlsb,
			$Script:Forest.Name, $htmlwhite
		)

		$first = 'Application partitions'
		If( $null -eq $AppPartitions )
		{
			$rowdata[ $rowIndex++ ] = @(
				$first, $htmlsb,
				'None', $htmlwhite
			)
		}
		Else
		{
			ForEach( $AppPartition in $AppPartitions )
			{
				$rowdata[ $rowIndex++ ] = @(
					$first,                   $htmlsb,
					$AppPartition.ToString(), $htmlwhite
				)
				$first = ''
			}
		}

		$first = 'Cross forest references'
		If( $null -eq $CrossForestReferences )
		{
			$rowdata[ $rowIndex++ ] = @(
				$first, $htmlsb,
				'None', $htmlwhite
			)
		}
		Else
		{
			ForEach( $CrossForestReference in $CrossForestReferences )
			{
				$rowdata[ $rowIndex++ ] = @(
					$first,                           $htmlsb,
					$CrossForestReference.ToString(), $htmlwhite
				)
				$first = ''
			}
		}

		$rowdata[ $rowIndex++ ] = @(
			'Domain naming master',            $htmlsb,
			$Script:Forest.DomainNamingMaster, $htmlwhite
		)

		$first = 'Domains in forest'
		If( $null -eq $Script:Domains )
		{
			$rowdata[ $rowIndex++ ] = @(
				$first, $htmlsb,
				'None', $htmlwhite
			)
		}
		Else
		{
			ForEach( $Domain in $tmpDomains2 )
			{
				$rowdata[ $rowIndex++ ] = @(
					$first,             $htmlsb,
					$Domain.ToString(), $htmlwhite
				)
				$first = ''
			}
		}

		$rowdata[ $rowIndex++ ] = @(
			'Partitions container',             $htmlsb,
			$Script:Forest.PartitionsContainer, $htmlwhite
		)

		$rowdata[ $rowIndex++ ] = @(
			'Root domain',            $htmlsb,
			$Script:ForestRootDomain, $htmlwhite
		)

		$rowdata[ $rowIndex++ ] = @(
			'Schema master',             $htmlsb,
			$Script:Forest.SchemaMaster, $htmlwhite
		)

		$first = 'Sites'
		If( $null -eq $Sites )
		{
			$rowdata[ $rowIndex++ ] = @(
				$first, $htmlsb,
				'None', $htmlwhite
			)
		}
		Else
		{
			ForEach( $Site in $Sites )
			{
				$rowdata[ $rowIndex++ ] = @(
					$first,           $htmlsb,
					$Site.ToString(), $htmlwhite
				)
				$first = ''
			}
		}

		$first = 'SPN suffixes'
		If( $null -eq $SPNSuffixes )
		{
			$rowdata[ $rowIndex++ ] = @(
				$first, $htmlsb,
				'None', $htmlwhite
			)
		}
		Else
		{
			ForEach( $SPNSuffix in $SPNSuffixes )
			{
				$rowdata[ $rowIndex++ ] = @(
					$first,                $htmlsb,
					$SPNSuffix.ToString(), $htmlwhite
				)
				$first = ''
			}
		}

		$rowdata[ $rowIndex++ ] = @(
			'Tombstone lifetime',      $htmlsb,
			"$TombstoneLifetime days", $htmlwhite
		)

		$first = 'UPN suffixes'
		If( $null -eq $UPNSuffixes )
		{
			$rowdata[ $rowIndex++ ] = @(
				$first, $htmlsb,
				'None', $htmlwhite
			)
		}
		Else
		{
			ForEach( $UPNSuffix in $UPNSuffixes )
			{
				$rowdata[ $rowIndex++ ] = @(
					$first,                $htmlsb,
					$UPNSuffix.ToString(), $htmlwhite
				)
				$first = ''
			}
		}

		$columnWidths  = @( '175x', '300px' )
		$columnHeaders = @(
			'Forest mode', $htmlsb,
			$ForestMode,   $htmlwhite
		)

		FormatHTMLTable -rowArray $rowdata `
			-columnArray $columnHeaders `
			-fixedWidth $columnWidths `
			-tablewidth '475'
		WriteHTMLLine 0 0 ''

		$rowData = $null
	}
}
#endregion

#region get all DCs in the forest
Function ProcessAllDCsInTheForest
{
	Function GetBasicDCInfo
	{
		Param
		(
			[Parameter( Mandatory = $true )]
			[String] $dn	## distinguishedName of a DC
		)

		#Write-Verbose "$(Get-Date -Format G): `t`t`t$dn"
		$DCName  = $dn.SubString( 0, $dn.IndexOf( '.' ) )
		$SrvName = $dn.SubString( $dn.IndexOf( '.' ) + 1 )
		## SrvName is actually the domain default naming context, e.g.,
		## DC=europe,DC=contoso,DC=com

		$Results = Get-ADDomainController -Identity $DCName -Server $SrvName -EA 0
		
		If($? -and $Null -ne $Results)
		{
			$GC       = $Results.IsGlobalCatalog.ToString()
			$ReadOnly = $Results.IsReadOnly.ToString()
			#ServerOS and ServerCore added in V2.20
			$ServerOS = $Results.OperatingSystem
			#https://blogs.msmvps.com/russel/2017/03/16/how-to-tell-if-youre-running-on-windows-server-core/
			$tmp = Get-RegistryValue "HKLM:\software\microsoft\windows nt\currentversion" "installationtype" $DCName
			If( $null -eq $tmp )
			{
				$ServerCore = 'Unknown'
			}
			ElseIf( $tmp -eq 'Server Core')
			{
				$ServerCore = 'Yes'
			}
			Else
			{
				$ServerCore = 'No'
			}
		}
		Else
		{
			$GC         = 'Unable to retrieve status'
			$ReadOnly   = $GC
			$ServerOS   = $GC
			$ServerCore = $GC
		}
		
		$hash = @{ 
			DCName     = $DC
			GC         = $GC
			ReadOnly   = $ReadOnly
			ServerOS   = $ServerOS
			ServerCore = $ServerCore
		}

		Return $hash
	} ## end Function GetBasicDCInfo

	Write-Verbose "$(Get-Date -Format G): `tDomain controllers"

	$txt = "Domain Controllers"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	#http://www.superedge.net/2012/09/how-to-get-ad-forest-in-powershell.html
	#http://msdn.microsoft.com/en-us/library/vstudio/system.directoryservices.activedirectory.forest.getforest%28v=vs.90%29
	#$ADContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("forest", $ADForest) 
	#2.16 change
	
	<# 3.06 change
	$ADContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("forest", $Script:Forest.Name)
	$Forest2 = [system.directoryservices.activedirectory.Forest]::GetForest($ADContext)
	Write-Verbose "$(Get-Date -Format G): `t`tBuild list of Domain controllers in the Forest"
	$AllDCs = @( $Forest2.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} )
	Write-Verbose "$(Get-Date -Format G): `t`tSort list of all Domain controllers"
	$AllDCs = @( $AllDCs | Sort-Object )
	$ADContext = $Null
	$Forest2 = $Null
	#>
	
	#new code added in 3.06 from MBS
	$Domains = $Script:Forest.domains
	$AllDCs = New-Object System.Collections.ArrayList
	
	ForEach($Domain in $Domains)
	{
		$sourceEntry        = New-Object System.Directoryservices.DirectoryEntry( "LDAP://$Domain" )
		$source             = New-Object System.Directoryservices.DirectorySearcher( $sourceEntry )
		$source.SearchScope = 'Subtree'
		$source.PageSize    = 1000

		$domainController = '805306369'
		$serverTrust      = ( 0x0080000 ).ToString()  # SERVER_TRUST_ACCOUNT
		$trustDelegate    = ( 0x0002000 ).ToString()  # TRUSTED_FOR_DELEGATION
		$partialSecrets   = ( 0x4000000 ).ToString()  # PARTIAL_SECRETS_ACCOUNT (RODC)
		$bitMask          = $serverTrust -bor $trustDelegate

		#$groupDC   = 516
		#$groupRODC = 521

		$source.Filter  = "(&(sAMAccountType=$domainController)(|(userAccountControl:1.2.840.113556.1.4.803:=$bitMask)(userAccountControl:1.2.840.113556.1.4.803:=$partialSecrets)))"

		$Results = $source.FindAll()

		ForEach($Result in $Results)
		{
			$null = $AllDCs.Add($Result.Properties.Item('dnshostname'))
		}

		If( $Results )
		{
			$Results.Dispose() 
		}
		
		$source.Dispose()
		$sourceEntry.Dispose()
	}

	$AllDCs = @( $AllDCs | Sort-Object )
	#end of new 3.06 code

	If($Null -eq $AllDCs)
	{
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "<None>"
		}
		If($Text)
		{
			Line 0 "<None>"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "None"
		}
	}
	Else
	{
		If($MSWORD -or $PDF)
		{
			[System.Collections.Hashtable[]] $WordTableRowHash = @();
			ForEach($DC in $AllDCs)
			{
				$WordTableRowHash += GetBasicDCInfo $DC
			}
		}
		If($Text)
		{
			# no longer works after 3.06 changes [int]$MaxDCNameLength = ($AllDCs | Measure-Object -Maximum -Property Length).Maximum
			#use this loop instead
			[int]$MaxDCNameLength = 0

			ForEach($item in $AllDCs)
			{
				If($item.Length -gt $MaxDCNameLength)
				{
					$MaxDCNameLength = $item.Length
				}
			}
			
			If($MaxDCNameLength -gt 4) #4 is length of "Name"
			{
				#2 is to allow for spacing between columns
				Line 1 ("Name" + (' ' * ($MaxDCNameLength - 2))) -NoNewLine
				Line 0 "Global Catalog  Read-only  Server OS                       Server Core"
				Line 1 ('=' * $MaxDCNameLength) -NoNewLine
				Line 0 "==============================================================================================="
			}
			Else
			{
				Line 1 "Name  Global Catalog  Read-only  Server OS                       Server Core"
				Line 1 "============================================================================"
			}
			
			ForEach($DC in $AllDCs)
			{
				$hash = GetBasicDCInfo $DC
				If(($DC).Length -lt ($MaxDCNameLength))
				{
					[int]$NumOfSpaces = ($MaxDCNameLength * -1) 
				}
				Else
				{
					[int]$NumOfSpaces = -4
				}
				Line 1 ( "{0,$NumOfSpaces}  {1,-15} {2,-10} {3,-31} {4,-3}" -f $DC,$hash.GC,$hash.Readonly,$hash.ServerOS,$hash.ServerCore)

				$Results = $Null
			}
			Line 0 ""
		}
		If($HTML)
		{
			#V3.00 pre-allocate rowdata
			## $rowdata = @()
			$rowData = New-Object System.Array[] $AllDCs.Count
			$rowIndx = 0
			
			ForEach($DC in $AllDCs)
			{
				$hash = GetBasicDCInfo $DC

				$rowdata[ $rowIndx ] = @(
					$DC,              $htmlwhite,
					$hash.GC,         $htmlwhite,
					$hash.ReadOnly,   $htmlwhite,
					$hash.ServerOS,   $htmlwhite,
					$hash.ServerCore, $htmlwhite
				)
				$rowIndx++
			}
		}
	}

	If($MSWord -or $PDF)
	{
		If($WordTableRowHash.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $WordTableRowHash `
			-Columns DCName, GC, ReadOnly, ServerOS, ServerCore `
			-Headers "Name", "Global Catalog", "Read-only", "Server OS", "Server Core" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

			$Table.Columns.Item(1).Width = 150
			$Table.Columns.Item(2).Width = 50
			$Table.Columns.Item(3).Width = 50
			$Table.Columns.Item(4).Width = 130
			$Table.Columns.Item(5).Width = 45

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($Text)
	{
		#nothing to do
	}
	If($HTML)
	{
		$columnHeaders = @(
			'Name',           $htmlsb,
			'Global Catalog', $htmlsb,
			'Read-only',      $htmlsb,
			'Server OS',      $htmlsb,
			'Server Core',    $htmlsb
		)

		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ''

		$rowdata = $null
	}

	$AllDCs = $Null
} ## end Function ProcessAllDCsInTheForest
#endregion

#region process CA information
Function ProcessCAInformation
{
	Write-Verbose "$(Get-Date -Format G): `tCA Information"
	
	$txt = "Certificate Authority Information"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	$txt = "Certification Authority Root(s)"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 4 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 $txt
	}

	$rootDSE = [ADSI]'LDAP://RootDSE'

	$configNC = $rootDSE.Properties[ 'configurationNamingContext' ].Value -as [String]

	$rootCA  = 'CN=Certification Authorities,CN=Public Key Services,CN=Services,' + $configNC
	$rootObj = [ADSI]( 'LDAP://' + $rootCA )
	$RootCnt = 0
	$AllCnt  = 0
	
	If($Null -ne $rootObj)
	{
		ForEach($obj in $rootObj.psbase.children)
		{
			$RootCnt++
			If($MSWORD -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Common name"; Value = $obj.cn; }
				$ScriptInformation += @{ Data = "Distinguished name"; Value = $obj.distinguishedName; }
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 70;
				$Table.Columns.Item(2).Width = 355;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Common name`t`t: " $obj.cn
				Line 1 "Distinguished name`t: " $obj.distinguishedName
				Line 1 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Common name",$htmlsb,$obj.cn,$htmlwhite)
				$rowdata += @(,('Distinguished name',$htmlsb,$obj.distinguishedName,$htmlwhite))
				$columnWidths = @("125","400")
				FormatHTMLTable -rowArray $rowdata `
				-columnArray $columnHeaders `
				-fixedWidth $columnWidths `
				-tablewidth "525"
				WriteHTMLLine 0 0 ' '

				$rowdata = $null
			}
		}
		
		If($RootCnt -gt 1)
		{
			$txt = "Possible error: There are more than one Certification Authority Root(s)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 $txt
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
				WriteHTMLLine 0 0 ''
			}
		}
	}
	#ElseIf($Null -eq $rootObj) changed in V2.22 by Michael B. Smith
	If($RootCnt -eq 0 -or $Null -eq $rootObj)
	{
		$txt = "No Certification Authority Root(s) were retrieved"
		#Write-Warning $txt
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt ''
			#WriteHTMLLine 0 0 ' '
		}
	}

	$txt = "Certification Authority Issuer(s)"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 4 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 $txt
	}

	$allCA = 'CN=Enrollment Services,CN=Public Key Services,CN=Services,' + $configNC
	$allObj = [ADSI]( 'LDAP://' + $allCA )
	
	If([string]::isnullorempty($allObj.psbase.children) -and !([string]::isnullorempty($rootObj.psbase.children)))
	{
		#uh oh error
		$txt = "Error: Certification Authority Root(s) exist, but no Certification Authority Issuers(s) (also known as Enrollment Agents) exist"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt
			WriteWordLine 0 0 ''
		}
		If($Text)
		{
			Line 0 $txt
			Line 0 ''
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt
			WriteHTMLLine 0 0 ''
		}
	}
	ElseIf(!([string]::isnullorempty($allObj.psbase.children)) -and !([string]::isnullorempty($rootObj.psbase.children)))
	{
		$AllCnt = 0
		ForEach($obj in $allObj.psbase.children)
		{
			$AllCnt++
			If($MSWORD -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Common name"; Value = $obj.cn; }
				$ScriptInformation += @{ Data = "Distinguished name"; Value = $obj.distinguishedName; }
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 70;
				$Table.Columns.Item(2).Width = 355;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ''
			}
			If($Text)
			{
				Line 1 "Common name`t`t: " $obj.cn
				Line 1 "Distinguished name`t: " $obj.distinguishedName
				Line 1 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Common name",$htmlsb,$obj.cn,$htmlwhite)
				$rowdata += @(,('Distinguished name',$htmlsb,$obj.distinguishedName,$htmlwhite))
				$columnWidths = @("125","400")
				FormatHTMLTable -rowArray $rowdata `
				-columnArray $columnHeaders `
				-fixedWidth $columnWidths `
				-tablewidth "525"
				WriteHTMLLine 0 0 ''

				$rowdata = $null
			}
		}

		If($AllCnt -lt $RootCnt)
		{
			$txt = "Error: More Certification Authority Root(s) exist than there are Certification Authority Issuers(s) (also known as Enrollment Agents)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt
				WriteWordLine 0 0 ''
			}
			If($Text)
			{
				Line 0 $txt
				Line 0 ''
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
				WriteHTMLLine 0 0 ''
			}
		}
	}
	ElseIf(([string]::isnullorempty($allObj.psbase.children)) -and ([string]::isnullorempty($rootObj.psbase.children)))
	{
		$txt = "No Certification Authority Issuer(s) were retrieved"
		#Write-Warning $txt
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt '' $Null 0 $False $True
			WriteWordLine 0 0 ''
		}
		If($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt ''
			WriteHTMLLine 0 0 ' '
		}
	}
	
	#if you have enrollment authorities and no roots – that’s a BIG error
	If($AllCnt -gt 0 -and $RootCnt -eq 0)
	{
		$txt = "Error: Certification Authority Issuers(s) (also known as Enrollment Agents) exist, but no Certification Authority Root(s) exist"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt
			WriteWordLine 0 0 ''
		}
		If($Text)
		{
			Line 0 $txt
			Line 0 ''
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt
			WriteHTMLLine 0 0 ''
		}
	}
} ## end Function ProcessCAInformation
#endregion

#region process ad optional features
Function ProcessADOptionalFeatures
{
	Write-Verbose "$(Get-Date -Format G): `tAD Optional Features"
	
	$txt = "AD Optional Features"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	$ADOptionalFeatures = Get-ADOptionalFeature -Filter * -EA 0
	
	If($? -and $Null -ne $ADOptionalFeatures)
	{
		ForEach($Item in $ADOptionalFeatures)
		{
			$Enabled       = 'No'
			$EnabledScopes = $null

			If($Item.EnabledScopes.Count -gt 0)
			{
				$Enabled = 'Yes'
				$EnabledScopes = $Item.EnabledScopes | Sort-Object 
			}
			
			If($MSWORD -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Feature name"; Value = $Item.Name; }
				$ScriptInformation += @{ Data = "Enabled"; Value = $Enabled; }
				
				If($Enabled -eq "Yes")
				{
					
					$cnt = 0
					ForEach($Scope in $EnabledScopes)
					{
						$cnt++
					
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Enabled Scopes"; Value = $Scope; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = $($Scope); }
						}
					}
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 125;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Feature Name`t: " $Item.Name
				Line 1 "Enabled`t`t: " $Enabled
				
				If($Enabled -eq "Yes")
				{
					Line 1 "Enabled Scopes`t: " -NoNewLine
					
					$cnt = 0
					ForEach($Scope in $EnabledScopes)
					{
						$cnt++
					
						If($cnt -eq 1)
						{
							Line 0 $Scope
						}
						Else
						{
							Line 3 "  $($Scope)"
						}
					}
					Line 0 ""
				}
				Else
				{
					Line 0 ""
				}
			}
			If($HTML)
			{
				#V3.00 - pre-allocate rowdata
				## $rowdata = @()
				$rowsRequired = 1
				If( $Enabled -eq 'Yes' -and $EnabledScopes )
				{
					If( $EnabledScopes -is [Array] )
					{
						$rowsRequired += $EnabledScopes.Count
					}
					Else
					{
						$rowsRequired++
					}
				}

				$rowdata  = New-Object System.Array[] $rowsRequired
				$rowIndex = 0

				$rowdata[ $rowIndex++ ] = @(
					'Enabled', $htmlsb,
					$Enabled,  $htmlwhite
				)

				If( $Enabled -eq 'Yes' )
				{
					$first = 'Enabled Scopes'
					ForEach( $Scope in $EnabledScopes )
					{
						$rowdata[ $rowIndex++ ] = @(
							$first, $htmlsb,
							$Scope, $htmlwhite
						)
						$first = ''
					}
				}

				$columnWidths  = @( '125px', '400px' )
				$columnHeaders = @(
					'Feature Name', $htmlsb,
					$Item.Name,     $htmlwhite
				)

				FormatHTMLTable -rowArray $rowdata `
				-columnArray $columnHeaders `
				-fixedWidth $columnWidths `
				-tablewidth '525'
				WriteHTMLLine 0 0 ''

				$rowdata = $null
			}
		}
	}
	ElseIf($? -and $Null -eq $ADOptionalFeatures)
	{
		$txt = "No AD Optional Features were retrieved"
		#Write-Warning $txt
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt -options $htmlBold
			WriteHTMLLine 0 0 ''
		}
	}
	Else
	{
		$txt = "Error retrieving AD Optional Features"
		Write-Warning $txt
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt -options $htmlBold
			WriteHTMLLine 0 0 ''
		}
	}
}
#endregion

#region process ad schema items

#new for 2.16
Function ProcessADSchemaItems
{
	Param(
		[String []] $Name = @( 
		'User-Account-Control', #Flags that control the behavior of a user account
		'msNPAllowDialin', #RAS Server
		'ms-Mcs-AdmPwd', #LAPS
		'ms-Mcs-AdmPwdExpirationTime', #LAPS
		'ms-SMS-Assignment-Site-Code', #SCCM
		'ms-SMS-Capabilities', #SCCM
		'ms-RTC-SIP-PoolAddress' #updated V3.00 for Lync/SfB
		'ms-RTC-SIP-DomainName' #updated V3.00 for Lync/SfB
		'ms-exch-schema-version-pt' #Exchange
		)
	)

	Write-Verbose "$(Get-Date -Format G): `tAD Schema Items"
	
	$txt = "AD Schema Items"
	$txt1 = "Just because a schema extension is Present does not mean that the product is in use." #V3.04 change wording
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 $txt
		WriteWordLine 0 0 $txt1 "" $Null 8 $False $True	
	}
	If($Text)
	{
		Line 0 $txt
		Line 0 $txt1
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 $txt
		WriteHTMLLine 0 0 $txt1
	}

	$rootDS   = [ADSI]'LDAP://RootDSE'
	$schemaNC = $rootDS.schemaNamingContext.Item( 0 )

	$SchemaItems = New-Object System.Collections.ArrayList
	ForEach( $item in $Name )
	{
		Write-Verbose "$(Get-Date -Format G): `t`tChecking for $item declared in schema"

		$objDN = 'LDAP://' + 'CN=' + $item + ',' + $schemaNC

		$obj = [ADSI] $objDN
		$mem = Get-Member -Name name -InputObject $obj

		$Itemobj = New-Object -TypeName PSObject

		Switch ($item)
		{
			'User-Account-Control'			{$tmp = "Flags that control the behavior of a user account";Break}
			'msNPAllowDialin'				{$tmp = "NPS/RAS Server";Break} #renamed in 3.05
			'ms-Mcs-AdmPwd'					{$tmp = "On-premises LAPS";Break} #renamed in 3.05
			'ms-Mcs-AdmPwdExpirationTime'	{$tmp = "On-premises LAPS";Break} #renamed in 3.05
			'ms-SMS-Assignment-Site-Code'	{$tmp = "MECM/SCCM";Break} #renamed in 3.05
			'ms-SMS-Capabilities'			{$tmp = "MECM/SCCM";Break} #renamed in 3.05
			'ms-RTC-SIP-PoolAddress'		{$tmp = "On-premises Lync/Skype for Business";Break} #V3.00 #renamed in 3.05
			'ms-RTC-SIP-DomainName'			{$tmp = "On-premises Lync/Skype for Business";Break} #V3.00 #renamed in 3.05
			'ms-exch-schema-version-pt' 	{$tmp = "On-premises Exchange";Break} #renamed in 3.05
			Default							{$tmp = "Unknown";Break}
		}
		
		$Itemobj | Add-Member -MemberType NoteProperty -Name ItemName	-Value $item
		$Itemobj | Add-Member -MemberType NoteProperty -Name ItemDesc	-Value $tmp
		If( $mem )
		{
			$Itemobj | Add-Member -MemberType NoteProperty -Name ItemState	-Value "Present"
		}
		Else
		{
			$Itemobj | Add-Member -MemberType NoteProperty -Name ItemState	-Value "Not Present"
		}
		$SchemaItems.Add($Itemobj) > $Null

		$mem = $null
		$obj = $null
	}

	If($MSWORD -or $PDF)
	{
		$ItemsWordTable = New-Object System.Collections.ArrayList
	}
	If($Text)
	{
		Line 1 "Schema item name                Present      Used for                                         "
		Line 1 "=============================================================================================="
	}
	If($HTML)
	{
		#V3.00 - pre-allocate rowData
		##$rowdata = @()
		$rowData = New-Object System.Array[] $SchemaItems.Count
		$rowIndx = 0
	}
	
	ForEach($item in $SchemaItems)
	{
		If($MSWORD -or $PDF)
		{
			$WordTableRowHash = @{ 
			ItemName = $Item.ItemName; 
			ItemState = $Item.ItemState;
			ItemDesc = $Item.ItemDesc
			}

			## Add the hash to the array
			$ItemsWordTable.Add($WordTableRowHash) > $Null
		}
		If($Text)
		{
			Line 1 ( "{0,-30}  {1,-11}  {2,-50}" -f $Item.ItemName,$Item.ItemState,$Item.ItemDesc)
		}
		If($HTML)
		{
			$rowdata[ $rowIndx++ ] = @(
				$Item.ItemName,  $htmlwhite,
				$Item.ItemState, $htmlwhite,
				$Item.ItemDesc,  $htmlwhite
			)
		}
	}

	If($MSWORD -or $PDF)
	{
		## Add the table to the document, using the hashtable
		If($ItemsWordTable.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ItemName, ItemState, ItemDesc `
			-Headers "Schema item name", "Present", "Used for" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			## IB - Set the header row format after the SetWordTableAlternateRowColor Function as it will paint the header row!
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150
			$Table.Columns.Item(2).Width = 75
			$Table.Columns.Item(3).Width = 200

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		$columnWidths  = @( '250px', '75px', '350px' )
		$columnHeaders = @(
			'Schema item name', $htmlsb,
			'Present',          $htmlsb,
			'Used for',         $htmlsb
		)

		FormatHTMLTable -rowArray $rowdata `
		-columnArray $columnHeaders `
		-fixedWidth $columnWidths `
		-tablewidth '675'
		WriteHTMLLine 0 0 ''

		$rowData = $null
	}
	
	$rootDS      = $null
	$schemaNC    = $null
	$objDN       = $null
	$SchemaItems = $null
} ## end Function ProcessADSchemaItems
#endregion

#region Site information
Function GetSiteLinkOptionText
{
	Param
	(
		$siteLinkOption
	)

	## https://msdn.microsoft.com/en-us/library/cc223552.aspx

	If( [String]::IsNullOrEmpty( $siteLinkOption ) )
	{
		Return 'Change Notification is Disabled'
	}

	Switch( $siteLinkOption )
	{
		'0'
			{
				Return 'Change Notification is Disabled'
			}
		'1'
			{
				Return 'Change Notification is Enabled with Compression'
			}
		'2'
			{
				Return 'Force sync in opposite direction at end of sync'
			}
		'3'
			{
				Return 'Change Notification is Enabled with Compression and Force sync in opposite direction at end of sync'
			}
		'4'
			{
				Return 'Disable compression of Change Notification messages'
			}
		'5'
			{
				Return 'Change Notification is Enabled without Compression'
			}
		'6'
			{
				Return 'Force sync in opposite direction at end of sync and Disable compression of Change Notification messages'
			}
		'7'
			{
				Return 'Change Notification is Enabled without Compression and Force sync in opposite direction at end of sync'
			}
		Default
			{
				Return "Unknown siteLink option: $siteLinkOption"
			}
	}

	## can't get here

} ## end Function

Function ProcessSiteInformation
{
	Write-Verbose "$(Get-Date -Format G): Writing sites and services data"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Sites and Services"
	}
	If($Text)
	{
		Line 0 "///  Sites and Services  \\\"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Sites and Services&nbsp;&nbsp;\\\"
	}
	
	#get site information

	$siteContainerDN = ("CN=Sites," + $Script:ConfigNC)

	$tmp = $Script:Forest.PartitionsContainer
	$ConfigurationBase = $tmp.SubString($tmp.IndexOf(",") + 1)
	$Sites = $Null
	$Sites = Get-ADObject -Filter 'ObjectClass -eq "site"' -SearchBase $ConfigurationBase -Properties Name, SiteObjectBl -Server $ADForest -EA 0 | Sort-Object Name

	If($? -and $Null -ne $Sites)
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing Inter-Site Transports"
		$AllSiteLinks = Get-ADObject -Searchbase $Script:ConfigNC -Server $ADForest `
		-Filter 'objectClass -eq "siteLink"' -Property Description, Options, Cost, ReplInterval, SiteList, Schedule -EA 0 `
		| Select-Object Name, Description, @{Name="SiteCount";Expression={$_.SiteList.Count}}, Cost, ReplInterval, `
		@{Name="Schedule";Expression={If($_.Schedule){If(($_.Schedule -Join " ").Contains("240")){"NonDefault"}Else{"24x7"}}Else{"24x7"}}}, `
		Options, SiteList, DistinguishedName

		If($MSWORD -or $PDF)
		{
			WriteWordLine 2 0 "Inter-Site Transports"
		}
		If($Text)
		{
			Line 0 "///  Inter-Site Transports  \\\"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "///&nbsp;&nbsp;Inter-Site Transports&nbsp;&nbsp;\\\"
		}

		#adapted from code provided by Goatee PFE
		#http://blogs.technet.com/b/ashleymcglone/archive/2011/06/29/report-and-edit-ad-site-links-from-powershell-turbo-your-ad-replication.aspx
		# Report of all site links and related settings
		
		If($? -and $Null -ne $AllSiteLinks)
		{
			ForEach($SiteLink in $AllSiteLinks)
			{
				Write-Verbose "$(Get-Date -Format G): `t`tProcessing site link $($SiteLink.Name)"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$SiteLinkTypeDN = @()
				$SiteLinkTypeDN = $SiteLink.DistinguishedName.Split(",")
				$SiteLinkType = $SiteLinkTypeDN[1].SubString(3)
				$SitesInLink = New-Object System.Collections.ArrayList
				$SiteLinkSiteList = $SiteLink.SiteList
				ForEach($xSite in $SiteLinkSiteList)
				{
					$tmp = $xSite.Split(",")
					$SitesInLink.Add("$($tmp[0].SubString(3))") > $Null
				}
				
				If($MSWord -or $PDF)
				{
					$ScriptInformation += @{ Data = "Name"; Value = $SiteLink.Name; }
					If(![String]::IsNullOrEmpty($SiteLink.Description))
					{
						$ScriptInformation += @{ Data = "Description"; Value = $SiteLink.Description; }
					}
					If($SitesInLink -ne "")
					{
						$cnt = 0
						ForEach($xSite in $SitesInLink)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								$ScriptInformation += @{ Data = "Sites in Link"; Value = $xSite; }
							}
							Else
							{
								$ScriptInformation += @{ Data = ""; Value = $xSite; }
							}
						}
					}
					Else
					{
						$ScriptInformation += @{ Data = "Sites in Link"; Value = "<None>"; }
					}
					$ScriptInformation += @{ Data = "Cost"; Value = $SiteLink.Cost.ToString(); }
					$ScriptInformation += @{ Data = "Replication Interval"; Value = $SiteLink.ReplInterval.ToString(); }
					$ScriptInformation += @{ Data = "Schedule"; Value = $SiteLink.Schedule; }

					$tmp = GetSiteLinkOptionText $SiteLink.Options
					$ScriptInformation += @{ Data = "Options"; Value = $tmp; }
					$ScriptInformation += @{ Data = "Type"; Value = $SiteLinkType; }

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitContent;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 0 "Name`t`t`t: " $SiteLink.Name
					If(![String]::IsNullOrEmpty($SiteLink.Description))
					{
						Line 0 "Description`t`t: " $SiteLink.Description
					}
					Line 0 "Sites in Link`t`t: " -NoNewLine
					If($SitesInLink -ne "")
					{
						$cnt = 0
						ForEach($xSite in $SitesInLink)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								Line 0 $xSite
							}
							Else
							{
								Line 3 "  $($xSite)"
							}
						}
					}
					Else
					{
						Line 0 "<None>"
					}
					Line 0 "Cost`t`t`t: " $SiteLink.Cost.ToString()
					Line 0 "Replication Interval`t: " $SiteLink.ReplInterval.ToString()
					Line 0 "Schedule`t`t: " $SiteLink.Schedule
					Line 0 "Options`t`t`t: " -NoNewLine

					$tmp = GetSiteLinkOptionText $SiteLink.Options
					Line 0 $tmp

					Line 0 "Type`t`t`t: " $SiteLinkType
					Line 0 ""
				}
				If($HTML)
				{
					$columnHeaders = @(
						'Name',         $htmlsb,
						$SiteLink.Name, $htmlwhite
					)
					$rowsRequired = 0
					$slDesc = ''
					If( ![String]::IsNullOrEmpty( $SiteLink.Description ) )
					{
						$slDesc = $SiteLink.Description
						$rowsRequired = 1
					}
					if( $SitesInLink -eq '' )
					{
						$rowsRequired++
					}
					else
					{
						$rowsRequired += $SitesInLink.Count	
					}
					$rowsRequired += 5

					$rowdata = New-Object System.Array[] $rowsRequired
					$rowIndx = 0

					if( $slDesc.Length -gt 0 )
					{
						$rowdata[ $rowIndx++ ] = @(
							'Description', $htmlsb, 
							$slDesc,       $htmlwhite
						)
					}

					If( $SitesInLink -eq '' )
					{
						$rowdata[ $rowIndx++ ] = @(
							'Sites in Link', $htmlsb,
							'None',          $htmlwhite
						)
					}
					Else
					{
						$cnt = 0
						ForEach( $xSite in $SitesInLink )
						{
							$cnt++
							
							If( $cnt -eq 1 )
							{
								$rowdata[ $rowIndx++ ] = @(
									'Sites in Link', $htmlsb,
									$xSite,          $htmlwhite
								)
							}
							Else
							{
								$rowdata[ $rowIndx++ ] = @(
									'',     $htmlsb,
									$xSite, $htmlwhite
								)
							}
						}
					}

					$rowdata[ $rowIndx++ ] = @(
						'Cost',                    $htmlsb,
						$SiteLink.Cost.ToString(), $htmlwhite
					)
					$rowdata[ $rowIndx++ ] = @(
						'Replication Interval',            $htmlsb,
						$SiteLink.ReplInterval.ToString(), $htmlwhite
					)
					$rowdata[ $rowIndx++ ] = @(
						'Schedule',         $htmlsb,
						$SiteLink.Schedule, $htmlwhite
					)

					$tmp = GetSiteLinkOptionText $SiteLink.Options
					$rowdata[ $rowIndx++ ] = @(
						'Options', $htmlsb,
						$tmp,      $htmlwhite
					)

					$rowdata[ $rowIndx++ ] = @(
						'Type',        $htmlsb,
						$SiteLinkType, $htmlwhite
					)

					$columnWidths = @( '125', '250' )

					FormatHTMLTable -rowArray $rowdata `
						-columnArray $columnHeaders `
						-fixedWidth $columnWidths `
						-tablewidth '375'

					WriteHTMLLine 0 0 ' '

					$rowdata = $null
				}
				$AllSiteLinks = $Null
			}
		}
		$AllSiteLinks = $Null
		
		ForEach($Site in $Sites)
		{
			Write-Verbose "$(Get-Date -Format G): `tProcessing site $($Site.Name)"
			If($MSWord -or $PDF)
			{
				WriteWordLine 2 0 "Site: " $Site.Name
				WriteWordLine 3 0 "Subnets"
			}
			If($Text)
			{
				Line 0 "///  Site: $($Site.Name)  \\\"
				Line 1 "Subnets"
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;Site: $($Site.Name)&nbsp;&nbsp;\\\"
				WriteHTMLLine 3 0 "Subnets"
			}
			
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing subnets"
			$subnetArray = New-Object -Type string[] -ArgumentList $Site.siteObjectBL.Count
			$i = 0
			$SitesiteObjectBL = $Site.siteObjectBL
			ForEach($subnetDN in $SitesiteObjectBL) 
			{
				$subnetName = $subnetDN.SubString(3, $subnetDN.IndexOf(",CN=Subnets,CN=Sites,") - 3)
				$subnetArray[$i] = $subnetName
				$i++
			}
			$subnetArray = @( $subnetArray | Sort-Object )
			#Fix in 3.03 to catch an empty array
			If($Null -eq $subnetArray -or $subnetArray.Count -eq 0)
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No Subnets linked to this site"
				}
				If($Text)
				{
					Line 2 "No Subnets linked to this site"
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No Subnets linked to this site"
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): `t`t`tSite $( $site.Name ) has $( $subnetArray.Count ) subnets"
				If($MSWord -or $PDF)
				{
					BuildMultiColumnTable $subnetArray
				}
				If($Text)
				{
					ForEach($xSubnet in $subnetArray)
					{
						Line 2 $xSubnet
					}
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = New-Object System.Array[] $subnetArray.Count
					$rowIndx = 0

					ForEach($xSubnet in $subnetArray)
					{
						$rowdata[ $rowIndx ] = @( $xSubnet, $htmlwhite )
						$rowIndx++
					}

					$columnHeaders = @( 'Subnets', $htmlsb )
					FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 ''

					$rowdata = $null
				}
			}
			
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing servers"
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Servers"
			}
			If($Text)
			{
				Line 1 "Servers"
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Servers"
			}
			$siteName = $Site.Name
			
			#build array of connect objects
			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing automatic connection objects"
			$Connections = New-Object System.Collections.ArrayList
			$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and options -bor 1' -Searchbase $Script:ConfigNC -Property DistinguishedName, fromServer -Server $ADForest -EA 0
			
			If($? -and $Null -ne $ConnectionObjects)
			{
				ForEach($ConnectionObject in $ConnectionObjects)
				{
					$xArray = $ConnectionObject.DistinguishedName.Split(",")
					#server name is 3rd item in array (element 2)
					$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
					$xArray = $ConnectionObject.FromServer.Split(",")
					#server name is 2nd item in array (element 1)
					$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
					#site name is 4th item in array (element 3)
					$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
					$xArray = $Null
					#V3.00 change to PSCustomObject
					
					$obj = [PSCustomObject] @{
						Name           = "<automatically generated>"						
						ToServer       = $ToServer						
						FromServer     = $FromServer						
						FromServerSite = $FromServerSite						
					}
					$Connections.Add($obj) > $Null
				}
			}
			
			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing manual connection objects"
			$ConnectionObjects = $Null
			$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and -not options -bor 1' -Searchbase $Script:ConfigNC -Property Name, DistinguishedName, fromServer -Server $ADForest -EA 0
			
			If($? -and $Null -ne $ConnectionObjects)
			{
				ForEach($ConnectionObject in $ConnectionObjects)
				{
					$xArray = $ConnectionObject.DistinguishedName.Split(",")
					#server name is 3rd item in array (element 2)
					$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
					$xArray = $ConnectionObject.FromServer.Split(",")
					#server name is 2nd item in array (element 1)
					$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
					#site name is 4th item in array (element 3)
					$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
					$xArray = $Null
					#V3.00 change to PSCustomObject
					
					$obj = [PSCustomObject] @{
						Name           = $ConnectionObject.Name						
						ToServer       = $ToServer						
						FromServer     = $FromServer						
						FromServerSite = $FromServerSite						
					}
					$Connections.Add($obj) > $Null
				}
			}

			If($Null -ne $Connections)
			{
				$Connections = $Connections | Sort-Object Name, ToServer, FromServer
			}
			
			#list each server
			$serverContainerDN = "CN=Servers,CN=" + $siteName + "," + $siteContainerDN
			$SiteServers = $Null
			$SiteServers = Get-ADObject -SearchBase $serverContainerDN -SearchScope OneLevel `
			-Filter { objectClass -eq "Server" } -Properties "DNSHostName" -Server $ADForest -EA 0 | `
			Select-Object DNSHostName, Name | Sort-Object DNSHostName
			
			If($? -and $Null -ne $SiteServers)
			{
				$First = $True
				ForEach($SiteServer in $SiteServers)
				{
					If(!$First)
					{
						If($MSword -or $PDF)
						{
							WriteWordLine 0 0 ""
						}
					}
					
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 0 $SiteServer.DNSHostName
						#for each server list each connection object
						If($Null -ne $Connections)
						{
							$Results = $Connections | Where-Object {$_.ToServer -eq $SiteServer.Name}

							If($? -and $Null -ne $Results)
							{
								WriteWordLine 0 0 "Connection Objects to source server $($SiteServer.Name)"
								[System.Collections.Hashtable[]] $WordTableRowHash = @();
								ForEach($Result in $Results)
								{
									$WordTableRowHash += @{ 
									ConnectionName = $Result.Name; 
									FromServer = $Result.FromServer; 
									FromSite = $Result.FromServerSite
									}
								}
								
								If($WordTableRowHash.Count -gt 0)	#3.07
								{
									$Table = AddWordTable -Hashtable $WordTableRowHash `
									-Columns ConnectionName, FromServer, FromSite `
									-Headers "Name", "From Server", "From Site" `
									-Format $wdTableGrid `
									-AutoFit $wdAutoFitFixed;

									SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

									$Table.Columns.Item(1).Width = 200;
									$Table.Columns.Item(2).Width = 100;
									$Table.Columns.Item(3).Width = 100;

									#indent the entire table 1 tab stop
									$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

									FindWordDocumentEnd
									$Table = $Null
									WriteWordLine 0 0 ""
								}
							}
						}
						Else
						{
							WriteWordLine 0 3 "Connection Objects: "
							WriteWordLine 0 4 "<None>"
							WriteWordLine 0 0 ""
						}
						$First = $False
					}
					If($Text)
					{
						Line 2 $SiteServer.DNSHostName
						Line 0 ""
						#for each server list each connection object
						If($Null -ne $Connections)
						{
							$Results = $Connections | Where-Object {$_.ToServer -eq $SiteServer.Name}

							If($? -and $Null -ne $Results)
							{
								Line 2 "Connection Objects to source server $($SiteServer.Name)"
								ForEach($Result in $Results)
								{
									Line 3 "Name`t`t: " $Result.Name
									Line 3 "From Server`t: " $Result.FromServer
									Line 3 "From Site`t: " $Result.FromServerSite
									Line 0 ""
								}
							}
						}
						Else
						{
							Line 3 "Connection Objects: <None>"
						}
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 $SiteServer.DNSHostName
						#WriteHTMLLine 0 0 " "
						#for each server list each connection object
						If($Null -ne $Connections)
						{
							$Results = $Connections | Where-Object {$_.ToServer -eq $SiteServer.Name}

							If($? -and $Null -ne $Results)
							{
								#V3.00 - pre-allocate rowdata
								## $rowdata = @()
								$rowCt   = 1
								If( $Results -is [Array] )
								{
									$rowCt = $Results.Count
								}
								$rowData = New-Object System.Array[] $rowCt
								$rowIndx = 0

								ForEach($Result in $Results)
								{
									#replace the <> characters since HTML doesn't like those in data
									$tmp = $Result.Name
									$tmp = $tmp.Replace( '>', '' )
									$tmp = $tmp.Replace( '<', '' )

									$rowdata[ $rowIndx ] = @(
										$tmp,$htmlwhite,
										$Result.FromServer,$htmlwhite,
										$Result.FromServerSite,$htmlwhite
									)
									$rowIndx++
								}

								$columnWidths  = @( '175px', '125px', '150px' )
								$columnHeaders = @(
									'Name',        $htmlsb,
									'From Server', $htmlsb,
									'From Site',   $htmlsb
								)
								$msg = "Connection Objects to source server $($SiteServer.Name)"
								FormatHTMLTable -tableHeader $msg `
									-rowArray $rowdata `
									-columnArray $columnHeaders `
									-fixedWidth $columnWidths `
									-tablewidth '450'
								WriteHTMLLine 0 0 ''

								$rowdata = $null
							}
						}
						Else
						{
							WriteHTMLLine 0 3 "Connection Objects: None"
						}
					}
				}
			}
			ElseIf(!$?)
			{
				Write-Warning "No Site Servers were retrieved."
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Warning: No Site Servers were retrieved" "" $Null 0 $False $True
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "Warning: No Site Servers were retrieved"
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Warning: No Site Servers were retrieved" -options $htmlBold
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No servers in this site"
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "No servers in this site"
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No servers in this site"
				}
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "No Sites were retrieved."
		$txt = "Warning: No Sites were retrieved"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
		}
		If($Text)
		{
			WriteWordLine 0 0 $txt
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt
		}
	}
	Else
	{
		$txt = "There were no sites found to retrieve"
		Write-Warning $txt
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
		}
		If($Text)
		{
			WriteWordLine 0 0 $txt
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt
		}
	}
} ## end Function ProcessSiteInformation
#endregion

#region domains
Function ProcessDomains
{
	Write-Verbose "$(Get-Date -Format G): Writing domain data"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Domain Information"
	}
	If($Text)
	{
		Line 0 "///  Domain Information  \\\"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Information&nbsp;&nbsp;\\\"
	}

	$Script:AllDomainControllers = New-Object System.Collections.ArrayList
	$First = $True

	#http://technet.microsoft.com/en-us/library/bb125224(v=exchg.150).aspx
	#http://support.microsoft.com/kb/556086/he
	#https://eightwone.com/references/ad-schema-versions/
	#https://eightwone.com/references/schema-versions/
	
	$SchemaVersionTable = @{ 
	"13" = "Windows 2000"; 
	"30" = "Windows 2003 RTM, SP1, SP2"; 
	"31" = "Windows 2003 R2";
	"44" = "Windows 2008"; 
	"47" = "Windows 2008 R2";
	"56" = "Windows Server 2012";
	"69" = "Windows Server 2012 R2";
	"72" = "Windows Server 2016 TP4";
	"87" = "Windows Server 2016";
	"88" = "Windows Server 2019/2022";	#added V2.20, updated in 2.22, updated in 3.10
	"4397" = "Exchange 2000 RTM"; 
	"4406" = "Exchange 2000 SP3";
	"6870" = "Exchange 2003 RTM, SP1, SP2"; 
	"6936" = "Exchange 2003 SP3"; 
	"10637" = "Exchange 2007 RTM";
	"11116" = "Exchange 2007 SP1"; 
	"14622" = "Exchange 2007 SP2, Exchange 2010 RTM";
	"14625" = "Exchange 2007 SP3";
	"14726" = "Exchange 2010 SP1";
	"14732" = "Exchange 2010 SP2";
	"14734" = "Exchange 2010 SP3";
	"15137" = "Exchange 2013 RTM";
	"15254" = "Exchange 2013 CU1";
	"15281" = "Exchange 2013 CU2";
	"15283" = "Exchange 2013 CU3";
	"15292" = "Exchange 2013 SP1/CU4";
	"15300" = "Exchange 2013 CU5";
	"15303" = "Exchange 2013 CU6";
	"15312" = "Exchange 2013 CU7 through CU23"; #updated in 2.20, updated in 2.24
	"15317" = "Exchange 2016 Preview and RTM"; #updated in 2.24
	"15323" = "Exchange 2016 CU1";
	"15325" = "Exchange 2016 CU2";
	"15326" = "Exchange 2016 CU3/CU4/CU5"; #added in 2.16
	"15330" = "Exchange 2016 CU6"; #added in 2.16
	"15332" = "Exchange 2016 CU7 through CU18"; #added in 2.16 and updated in 2.20, updated in 2.22, updated in 2.24, updated in 3.02
	"15333" = "Exchange 2016 CU19/CU20"; #added in 3.02, updated in 3.05
	"15334" = "Exchange 2016 CU21-CU23"; #added in 3.05, updated in 3.08, updated in 3.10
	"17000" = "Exchange 2019 RTM/CU1"; #added in 2.22, updated in 2.24
	"17001" = "Exchange 2019 CU2-CU7"; #added in 2.24, updated in 3.02
	"17002" = "Exchange 2019 CU8/CU9"; #added in 3.02, updated in 3.05
	"17003" = "Exchange 2019 CU10-CU12"; #added in 3.05, updated in 3.08, updated in 3.10
	}

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing domain $($Domain)"

		$DomainInfo = Get-ADDomain -Identity $Domain -EA 0

		If( !$? )
		{
			$txt = "Error retrieving domain data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If( $null -eq $DomainInfo )
		{
			$txt = "No Domain data was retrieved for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If($? -and $Null -ne $DomainInfo)
		{
			If(($MSWORD -or $PDF) -and !$First)
			{
				#put each domain, starting with the second, on a new page
				$Script:selection.InsertNewPage()
			}
			
			If($Domain -eq $Script:ForestRootDomain)
			{
				If($MSWORD -or $PDF)
				{
					WriteWordLine 2 0 "$($Domain) (Forest Root)"
				}
				If($Text)
				{
					Line 0 "///  $($Domain) (Forest Root)  \\\"
				}
				If($HTML)
				{
					WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($Domain) (Forest Root)&nbsp;&nbsp;\\\"
				}
			}
			Else
			{
				If($MSWORD -or $PDF)
				{
					WriteWordLine 2 0 $Domain
				}
				If($Text)
				{
					Line 0 "///  $($Domain)  \\\"
				}
				If($HTML)
				{
					WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($Domain)&nbsp;&nbsp;\\\"
				}
			}

			Switch ($DomainInfo.DomainMode)
			{
				"0"	{$DomainMode = "Windows 2000"; Break}
				"1" {$DomainMode = "Windows Server 2003 mixed"; Break}
				"2" {$DomainMode = "Windows Server 2003"; Break}
				"3" {$DomainMode = "Windows Server 2008"; Break}
				"4" {$DomainMode = "Windows Server 2008 R2"; Break}
				"5" {$DomainMode = "Windows Server 2012"; Break}
				"6" {$DomainMode = "Windows Server 2012 R2"; Break}
				"7" {$DomainMode = "Windows Server 2016"; Break}	#added V2.20
				"Windows2000Domain"   		{$DomainMode = "Windows 2000"; Break}
				"Windows2003Mixed"    		{$DomainMode = "Windows Server 2003 mixed"; Break}
				"Windows2003Domain"   		{$DomainMode = "Windows Server 2003"; Break}
				"Windows2008Domain"   		{$DomainMode = "Windows Server 2008"; Break}
				"Windows2008R2Domain" 		{$DomainMode = "Windows Server 2008 R2"; Break}
				"Windows2012Domain"   		{$DomainMode = "Windows Server 2012"; Break}
				"Windows2012R2Domain" 		{$DomainMode = "Windows Server 2012 R2"; Break}
				"WindowsThresholdDomain"	{$DomainMode = "Windows Server 2016 TP"; Break}
				"Windows2016Domain"			{$DomainMode = "Windows Server 2016"; Break}
				"UnknownDomain"       		{$DomainMode = "Unknown Domain Mode"; Break}
				Default               		{$DomainMode = "Unable to determine Domain Mode: $($DomainInfo.DomainMode)"; Break}
			}
			
			$ADSchemaInfo = $Null
			$ExchangeSchemaInfo = $Null
			
			$ADSchemaInfo = Get-ADObject (Get-ADRootDSE -Server $Domain -EA 0).schemaNamingContext `
			-Property objectVersion -Server $Domain -EA 0
			
			If($? -and $Null -ne $ADSchemaInfo)
			{
				$ADSchemaVersion = $ADSchemaInfo.objectversion
				$ADSchemaVersionName = $SchemaVersionTable.Get_Item("$ADSchemaVersion")
				If($Null -eq $ADSchemaVersionName)
				{
					$ADSchemaVersionName = "Unknown"
				}
			}
			Else
			{
				$ADSchemaVersion = "Unknown"
				$ADSchemaVersionName = "Unknown"
			}
			
			If($Domain -eq $Script:ForestRootDomain)
			{
				#try/catch and -LDAPFilter added in 3.03 to suppress the error 
				#Get-ADObject : Directory object not found
				#+ ... chemaInfo = Get-ADObject "cn=ms-exch-schema-version-pt,cn=Schema,cn=C ...				
				try
				{
					$ExchangeSchemaInfo = Get-ADObject -LDAPFilter "cn=ms-exch-schema-version-pt,cn=Schema,cn=Configuration,$($DomainInfo.DistinguishedName)" `
					-properties rangeupper -Server $Domain -EA 0
				}
				
				catch
				{
					#do nothing but suppress the error message if Exchange is not installed in the domain
				}

				If($? -and $Null -ne $ExchangeSchemaInfo)
				{
					$ExchangeSchemaVersion = $ExchangeSchemaInfo.rangeupper
					$ExchangeSchemaVersionName = $SchemaVersionTable.Get_Item("$ExchangeSchemaVersion")
					If($Null -eq $ExchangeSchemaVersionName)
					{
						$ExchangeSchemaVersionName = "Unknown"
					}
				}
				Else
				{
					$ExchangeSchemaVersion = "Unknown"
					$ExchangeSchemaVersionName = "Unknown"
				}
			}

			If($Null -eq $DomainInfo.LastLogonReplicationInterval)
			{
				$LastLogonReplicationInterval = "Default 1 day"
			}
			Else
			{
				$LastLogonReplicationInterval = $DomainInfo.LastLogonReplicationInterval.ToString()
			}
			
			#added in 3.09
			#https://www.jorgebernhardt.com/how-to-change-attribute-ms-ds-machineaccountquota/
			$tmpMachineAccountQuota = Get-ADObject -Identity $DomainInfo.DistinguishedName -Properties ms-DS-MachineAccountQuota -EA 0
			
			If($? -and $Null -ne $tmpMachineAccountQuota)
			{
				$MachineAccountQuota = $tmpMachineAccountQuota.'ms-DS-MachineAccountQuota'.ToString()
			}
			Else
			{
				$MachineAccountQuota = "Unable to retrieve"
			}
			$tmpMachineAccountQuota = $Null
			#end add
			
			If($MSWORD -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Domain mode"; Value = $DomainMode; }
				$ScriptInformation += @{ Data = "Domain name"; Value = $DomainInfo.Name; }
				$ScriptInformation += @{ Data = "NetBIOS name"; Value = $DomainInfo.NetBIOSName; }
				$ScriptInformation += @{ Data = "AD Schema"; Value = "($($ADSchemaVersion)) - $($ADSchemaVersionName)"; }
				$DNSSuffixes = $DomainInfo.AllowedDNSSuffixes | Sort-Object 
				If($Null -eq $DNSSuffixes)
				{
					$ScriptInformation += @{ Data = "Allowed DNS Suffixes"; Value = "<None>"; }
				}
				Else
				{
					$cnt = 0
					ForEach($DNSSuffix in $DNSSuffixes)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Allowed DNS Suffixes"; Value = "$($DNSSuffix.ToString())"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = "$($DNSSuffix.ToString())"; }
						}
					}
				}
				$ChildDomains = $DomainInfo.ChildDomains | Sort-Object 
				If($Null -eq $ChildDomains)
				{
					$ScriptInformation += @{ Data = "Child domains"; Value = "<None>"; }
				}
				Else
				{
					$cnt = 0 
					ForEach($ChildDomain in $ChildDomains)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Child domains"; Value = "$($ChildDomain.ToString())"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = "$($ChildDomain.ToString())"; }
						}
					}
				}
				$ScriptInformation += @{ Data = "Default computers container"; Value = $DomainInfo.ComputersContainer; }
				$ScriptInformation += @{ Data = "Default users container"; Value = $DomainInfo.UsersContainer; }
				$ScriptInformation += @{ Data = "Deleted objects container"; Value = $DomainInfo.DeletedObjectsContainer; }
				$ScriptInformation += @{ Data = "Distinguished name"; Value = $DomainInfo.DistinguishedName; }
				$ScriptInformation += @{ Data = "DNS root"; Value = $DomainInfo.DNSRoot; }
				$ScriptInformation += @{ Data = "Domain controllers container"; Value = $DomainInfo.DomainControllersContainer; }
				$ScriptInformation += @{ Data = "Domain SID"; Value = $DomainInfo.DomainSID; }
				If(![String]::IsNullOrEmpty($ExchangeSchemaInfo))
				{
					$ScriptInformation += @{ Data = "Exchange Schema"; Value = "($($ExchangeSchemaVersion)) - $($ExchangeSchemaVersionName)"; }
				}
				$ScriptInformation += @{ Data = "Foreign security principals container"; Value = $DomainInfo.ForeignSecurityPrincipalsContainer; }
				$ScriptInformation += @{ Data = "Infrastructure master"; Value = $DomainInfo.InfrastructureMaster; }
				$ScriptInformation += @{ Data = "Last logon replication interval"; Value = $LastLogonReplicationInterval; }
				$ScriptInformation += @{ Data = "Lost and Found container"; Value = $DomainInfo.LostAndFoundContainer; }
				$ScriptInformation += @{ Data = "Machine account quota"; Value = $MachineAccountQuota; }
				If(![String]::IsNullOrEmpty($DomainInfo.ManagedBy))
				{
					$ScriptInformation += @{ Data = "Managed by"; Value = $DomainInfo.ManagedBy; }
				}
				$ScriptInformation += @{ Data = "PDC Emulator"; Value = $DomainInfo.PDCEmulator; }
				If( ( validObject $DomainInfo PublicKeyRequiredPasswordRolling ) -and $null -ne $DomainInfo.PublicKeyRequiredPasswordRolling )
				{
					$ScriptInformation += @{ Data = "Public key required password rolling"; Value = $DomainInfo.PublicKeyRequiredPasswordRolling.ToString(); }
				}
				$ScriptInformation += @{ Data = "Quotas container"; Value = $DomainInfo.QuotasContainer; }
				$ReadOnlyReplicas = $DomainInfo.ReadOnlyReplicaDirectoryServers | Sort-Object 
				If($Null -eq $ReadOnlyReplicas)
				{
					$ScriptInformation += @{ Data = "Read-only replica directory servers"; Value = "<None>"; }
				}
				Else
				{
					$cnt = 0 
					ForEach($ReadOnlyReplica in $ReadOnlyReplicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Read-only replica directory servers"; Value = "$($ReadOnlyReplica.ToString())"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = "$($ReadOnlyReplica.ToString())"; }
						}
					}
				}
				$Replicas = $DomainInfo.ReplicaDirectoryServers | Sort-Object 
				If($Null -eq $Replicas)
				{
					$ScriptInformation += @{ Data = "Replica directory servers"; Value = "<None>"; }
				}
				Else
				{
					$cnt = 0 
					ForEach($Replica in $Replicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Replica directory servers"; Value = "$($Replica.ToString())"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = "$($Replica.ToString())"; }
						}
					}
				}
				$ScriptInformation += @{ Data = "RID Master"; Value = $DomainInfo.RIDMaster; }
				$SubordinateReferences = $DomainInfo.SubordinateReferences | Sort-Object 
				If($Null -eq $SubordinateReferences)
				{
					$ScriptInformation += @{ Data = "Subordinate references"; Value = "<None>"; }
				}
				Else
				{
					$cnt = 0
					ForEach($SubordinateReference in $SubordinateReferences)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Subordinate references"; Value = "$($SubordinateReference.ToString())"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = "$($SubordinateReference.ToString())"; }
						}
					}
				}
				$ScriptInformation += @{ Data = "Systems container"; Value = $DomainInfo.SystemsContainer; }
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

				$Table.Columns.Item(1).Width = 175
				$Table.Columns.Item(2).Width = 325

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Domain mode`t`t`t`t: " $DomainMode
				Line 1 "Domain name`t`t`t`t: " $DomainInfo.Name
				Line 1 "NetBIOS name`t`t`t`t: " $DomainInfo.NetBIOSName
				Line 1 "AD Schema`t`t`t`t: ($($ADSchemaVersion)) - $($ADSchemaVersionName)"
				Line 1 "Allowed DNS Suffixes`t`t`t: " -NoNewLine
				$DNSSuffixes = $DomainInfo.AllowedDNSSuffixes | Sort-Object 
				If($Null -eq $DNSSuffixes)
				{
					 Line 0 "<None>"
				}
				Else
				{
					$cnt = 0
					ForEach($DNSSuffix in $DNSSuffixes)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							Line 0 $DNSSuffix.ToString()
						}
						Else
						{
							Line 6 "  $($DNSSuffix.ToString())"
						}
					}
				}
				Line 1 "Child domains`t`t`t`t: " -NoNewLine
				$ChildDomains = $DomainInfo.ChildDomains | Sort-Object 
				If($Null -eq $ChildDomains)
				{
					Line 0 "<None>"
				}
				Else
				{
					$cnt = 0
					ForEach($ChildDomain in $ChildDomains)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							Line 0 $ChildDomain.ToString()
						}
						Else
						{
							Line 6 "  $($ChildDomain.ToString())"
						}
					}
				}
				Line 1 "Default computers container`t`t: " $DomainInfo.ComputersContainer
				Line 1 "Default users container`t`t`t: " $DomainInfo.UsersContainer
				Line 1 "Deleted objects container`t`t: " $DomainInfo.DeletedObjectsContainer
				Line 1 "Distinguished name`t`t`t: " $DomainInfo.DistinguishedName
				Line 1 "DNS root`t`t`t`t: " $DomainInfo.DNSRoot
				Line 1 "Domain controllers container`t`t: " $DomainInfo.DomainControllersContainer
				Line 1 "Domain SID`t`t`t`t: " $DomainInfo.DomainSID
				If(![String]::IsNullOrEmpty($ExchangeSchemaInfo))
				{
					Line 1 "Exchange Schema`t`t`t`t: ($($ExchangeSchemaVersion)) - $($ExchangeSchemaVersionName)"
				}
				Line 1 "Foreign security principals container`t: " $DomainInfo.ForeignSecurityPrincipalsContainer
				Line 1 "Infrastructure master`t`t`t: " $DomainInfo.InfrastructureMaster
				Line 1 "Last logon replication interval`t`t: " $LastLogonReplicationInterval
				Line 1 "Lost and Found container`t`t: " $DomainInfo.LostAndFoundContainer
				Line 1 "Machine account quota`t`t`t: " $MachineAccountQuota
				If(![String]::IsNullOrEmpty($DomainInfo.ManagedBy))
				{
					Line 1 "Managed by`t`t`t`t: " $DomainInfo.ManagedBy
				}
				Line 1 "PDC Emulator`t`t`t`t: " $DomainInfo.PDCEmulator
				If( ( validObject $DomainInfo PublicKeyRequiredPasswordRolling ) -and $null -ne $DomainInfo.PublicKeyRequiredPasswordRolling )
				{
					Line 1 "Public key required password rolling`t: " $DomainInfo.PublicKeyRequiredPasswordRolling.ToString()
				}
				Line 1 "Quotas container`t`t`t: " $DomainInfo.QuotasContainer
				Line 1 "Read-only replica directory servers`t: " -NoNewLine
				$ReadOnlyReplicas = $DomainInfo.ReadOnlyReplicaDirectoryServers | Sort-Object 
				If($Null -eq $ReadOnlyReplicas)
				{
					Line 0 "<None>"
				}
				Else
				{
					$cnt = 0
					ForEach($ReadOnlyReplica in $ReadOnlyReplicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							Line 0 $ReadOnlyReplica.ToString()
						}
						Else
						{
							Line 6 "  $($ReadOnlyReplica.ToString())"
						}
					}
				}
				Line 1 "Replica directory servers`t`t: " -NoNewLine
				$Replicas = $DomainInfo.ReplicaDirectoryServers | Sort-Object 
				If($Null -eq $Replicas)
				{
					Line 0 "<None>"
				}
				Else
				{
					$cnt = 0
					ForEach($Replica in $Replicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							Line 0 $Replica.ToString()
						}
						Else
						{
							Line 6 "  $($Replica.ToString())"
						}
					}
				}
				Line 1 "RID Master`t`t`t`t: " $DomainInfo.RIDMaster
				Line 1 "Subordinate references`t`t`t: " -NoNewLine
				$SubordinateReferences = $DomainInfo.SubordinateReferences | Sort-Object 
				If($Null -eq $SubordinateReferences)
				{
					Line 0 "<None>"
				}
				Else
				{
					$cnt = 0
					ForEach($SubordinateReference in $SubordinateReferences)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							Line 0 $SubordinateReference.ToString()
						}
						Else
						{
							Line 6 "  $($SubordinateReference.ToString())"
						}
					}
				}
				Line 1 "Systems container`t`t`t: " $DomainInfo.SystemsContainer
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()  #V3.00 - an ArrayList is probably a better option - but is a low-impact routine (only run once)
				$columnHeaders = @("Domain mode",$htmlsb,$DomainMode,$htmlwhite)
				$rowdata += @(,('Domain name',$htmlsb,$DomainInfo.Name,$htmlwhite))
				$rowdata += @(,('NetBIOS name',$htmlsb,$DomainInfo.NetBIOSName,$htmlwhite))
				$rowdata += @(,('AD Schema',$htmlsb,"($($ADSchemaVersion)) - $($ADSchemaVersionName)",$htmlwhite))
				$DNSSuffixes = $DomainInfo.AllowedDNSSuffixes | Sort-Object 
				If($Null -eq $DNSSuffixes)
				{
					$rowdata += @(,('Allowed DNS Suffixes',$htmlsb,"None",$htmlwhite))
				}
				Else
				{
					$cnt = 0
					ForEach($DNSSuffix in $DNSSuffixes)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$rowdata += @(,('Allowed DNS Suffixes',$htmlsb,"$($DNSSuffix.ToString())",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',$htmlsb,"$($DNSSuffix.ToString())",$htmlwhite))
						}
					}
				}
				$ChildDomains = $DomainInfo.ChildDomains | Sort-Object 
				If($Null -eq $ChildDomains)
				{
					$rowdata += @(,('Child domains',$htmlsb,"None",$htmlwhite))
				}
				Else
				{
					$cnt = 0 
					ForEach($ChildDomain in $ChildDomains)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$rowdata += @(,('Child domains',$htmlsb,"$($ChildDomain.ToString())",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',$htmlsb,"$($ChildDomain.ToString())",$htmlwhite))
						}
					}
				}
				$rowdata += @(,('Default computers container',$htmlsb,$DomainInfo.ComputersContainer,$htmlwhite))
				$rowdata += @(,('Default users container',$htmlsb,$DomainInfo.UsersContainer,$htmlwhite))
				$rowdata += @(,('Deleted objects container',$htmlsb,$DomainInfo.DeletedObjectsContainer,$htmlwhite))
				$rowdata += @(,('Distinguished name',$htmlsb,$DomainInfo.DistinguishedName,$htmlwhite))
				$rowdata += @(,('DNS root',$htmlsb,$DomainInfo.DNSRoot,$htmlwhite))
				$rowdata += @(,('Domain controllers container',$htmlsb,$DomainInfo.DomainControllersContainer,$htmlwhite))
				$rowdata += @(,('Domain SID',$htmlsb,$DomainInfo.DomainSID,$htmlwhite))
				If(![String]::IsNullOrEmpty($ExchangeSchemaInfo))
				{
					$rowdata += @(,('Exchange Schema',$htmlsb,"($($ExchangeSchemaVersion)) - $($ExchangeSchemaVersionName)",$htmlwhite))
				}
				$rowdata += @(,('Foreign security principals container',$htmlsb,$DomainInfo.ForeignSecurityPrincipalsContainer,$htmlwhite))
				$rowdata += @(,('Infrastructure master',$htmlsb,$DomainInfo.InfrastructureMaster,$htmlwhite))
				$rowdata += @(,("Last logon replication interval",$htmlsb,$LastLogonReplicationInterval,$htmlwhite))
				$rowdata += @(,('Lost and Found container',$htmlsb,$DomainInfo.LostAndFoundContainer,$htmlwhite))
				$rowdata += @(,("Machine account quota",$htmlsb,$MachineAccountQuota,$htmlwhite))
				If(![String]::IsNullOrEmpty($DomainInfo.ManagedBy))
				{
					$rowdata += @(,('Managed by',$htmlsb,$DomainInfo.ManagedBy,$htmlwhite))
				}
				$rowdata += @(,('PDC Emulator',$htmlsb,$DomainInfo.PDCEmulator,$htmlwhite))
				If( ( validObject $DomainInfo PublicKeyRequiredPasswordRolling ) -and $null -ne $DomainInfo.PublicKeyRequiredPasswordRolling )
				{
					$rowdata += @(,("Public key required password rolling",$htmlsb,$DomainInfo.PublicKeyRequiredPasswordRolling.ToString(),$htmlwhite))
				}
				$rowdata += @(,('Quotas container',$htmlsb,$DomainInfo.QuotasContainer,$htmlwhite))
				$ReadOnlyReplicas = $DomainInfo.ReadOnlyReplicaDirectoryServers | Sort-Object 
				If($Null -eq $ReadOnlyReplicas)
				{
					$rowdata += @(,('Read-only replica directory servers',$htmlsb,"None",$htmlwhite))
				}
				Else
				{
					$cnt = 0 
					ForEach($ReadOnlyReplica in $ReadOnlyReplicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$rowdata += @(,('Read-only replica directory servers',$htmlsb,"$($ReadOnlyReplica.ToString())",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',$htmlsb,"$($ReadOnlyReplica.ToString())",$htmlwhite))
						}
					}
				}
				$Replicas = $DomainInfo.ReplicaDirectoryServers | Sort-Object 
				If($Null -eq $Replicas)
				{
					$rowdata += @(,('Replica directory servers',$htmlsb,"None",$htmlwhite))
				}
				Else
				{
					$cnt = 0 
					ForEach($Replica in $Replicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$rowdata += @(,('Replica directory servers',$htmlsb,"$($Replica.ToString())",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',$htmlsb,"$($Replica.ToString())",$htmlwhite))
						}
					}
				}
				$rowdata += @(,('RID Master',$htmlsb,$DomainInfo.RIDMaster,$htmlwhite))
				$SubordinateReferences = $DomainInfo.SubordinateReferences | Sort-Object 
				If($Null -eq $SubordinateReferences)
				{
					$rowdata += @(,('Subordinate references',$htmlsb,"None",$htmlwhite))
				}
				Else
				{
					$cnt = 0
					ForEach($SubordinateReference in $SubordinateReferences)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$rowdata += @(,('Subordinate references',$htmlsb,"$($SubordinateReference.ToString())",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',$htmlsb,"$($SubordinateReference.ToString())",$htmlwhite))
						}
					}
				}
				$rowdata += @(,('Systems container',$htmlsb,$DomainInfo.SystemsContainer,$htmlwhite))
				
				$columnWidths = @("250","300")
				FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "550"
				#WriteHTMLLine 0 0 ' '

				$rowData = $null
			}

			Write-Verbose "$(Get-Date -Format G): `t`tGetting domain trusts"
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Domain Trusts"
			}
			If($Text)
			{
				Line 0 "Domain Trusts: "
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Domain Trusts"
			}
			
			$ADDomainTrusts = $Null
			$ADDomainTrusts = Get-ADObject -Filter {ObjectClass -eq "trustedDomain"} `
			-Server $Domain -Properties * -EA 0

			If($? -and $Null -ne $ADDomainTrusts)
			{
				
				ForEach($Trust in $ADDomainTrusts) 
				{
					If($MSWord -or $PDF)
					{
						[System.Collections.Hashtable[]] $ScriptInformation = @()
						$ScriptInformation += @{ Data = "Name"; Value = $Trust.Name; }
						
						If(![String]::IsNullOrEmpty($Trust.Description))
						{
							$ScriptInformation += @{ Data = "Description"; Value = $Trust.Description; }
						}
						
						$ScriptInformation += @{ Data = "Created"; Value = $Trust.Created; }
						$ScriptInformation += @{ Data = "Modified"; Value = $Trust.Modified; }

						$TrustExtendedAttributes = Get-ADTrustInfo $Trust
						
						$ScriptInformation += @{ Data = "Type"; Value = $TrustExtendedAttributes.TrustType; }

						$cnt = 0
						ForEach($attribute in $TrustExtendedAttributes.TrustAttribute)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								$ScriptInformation += @{ Data = "Attributes"; Value = $attribute.ToString(); }
							}
							Else
							{
								$ScriptInformation += @{ Data = ""; Value = "$($attribute.ToString())"; }
							}
						}

						$ScriptInformation += @{ Data = "Direction"; Value = $TrustExtendedAttributes.TrustDirection; }

						$Table = AddWordTable -Hashtable $ScriptInformation `
						-Columns Data,Value `
						-List `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 175;
						$Table.Columns.Item(2).Width = 300;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
					If($Text)
					{
						Line 1 "Name`t`t: " $Trust.Name 
						
						If(![String]::IsNullOrEmpty($Trust.Description))
						{
							Line 1 "Description`t: " $Trust.Description
						}
						
						Line 1 "Created`t`t: " $Trust.Created
						Line 1 "Modified`t: " $Trust.Modified

						$TrustExtendedAttributes = Get-ADTrustInfo $Trust
						
						Line 1 "Type`t`t: " $TrustExtendedAttributes.TrustType
						Line 1 "Attributes`t: " -NoNewLine
						$cnt = 0
						ForEach($attribute in $TrustExtendedAttributes.Trustattribute)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								Line 0 $attribute.ToString()
							}
							Else
							{
								Line 3 "  $($attribute.ToString())"
							}
						}

						Line 1 "Direction`t: " $TrustExtendedAttributes.TrustDirection
						Line 0 ""
					}
					If($HTML)
					{
						$rowdata = @()
						$columnHeaders = @("Name",$htmlsb,$Trust.Name,$htmlwhite)
						
						If(![String]::IsNullOrEmpty($Trust.Description))
						{
							$rowdata += @(,('Description',$htmlsb,$Trust.Description,$htmlwhite))
						}
						
						$rowdata += @(,('Created',$htmlsb,$Trust.Created,$htmlwhite))
						$rowdata += @(,('Modified',$htmlsb,$Trust.Modified,$htmlwhite))
	
						$TrustExtendedAttributes = Get-ADTrustInfo $Trust
						 
						$rowdata += @(,('Type',$htmlsb,$TrustExtendedAttributes.TrustType,$htmlwhite))

						
						$cnt = 0
						ForEach($attribute in $TrustExtendedAttributes.Trustattribute)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								$rowdata += @(,('Attributes',$htmlsb,$attribute.ToString(),$htmlwhite))
							}
							Else
							{
								$rowdata += @(,('',$htmlsb,$attribute.ToString(),$htmlwhite))
							}
						}

						$rowdata += @(,('Direction',$htmlsb,$TrustExtendedAttributes.TrustDirection,$htmlwhite))

						$columnWidths = @("175","300")
						FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "475"
						WriteHTMLLine 0 0 ''

						$rowData = $null
					}
				}
			}
			ElseIf(!$?)
			{
				#error retrieving domain trusts
				Write-Warning "Error retrieving domain trusts for $($Domain)"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Error retrieving domain trusts for $($Domain)" "" $Null 0 $False $True
				}
				If($Text)
				{
					Line 0 "Error retrieving domain trusts for $($Domain)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Error retrieving domain trusts for $($Domain)" -options $htmlBold
				}
			}
			Else
			{
				#no domain trust data
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "<None>"
				}
				If($Text)
				{
					Line 1 "<None>"
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "None"
				}
			}
			
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing domain controllers"
			#3.11 move the headings out so the error or notice hss a section heading
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Domain Controllers"
			}
			If($Text)
			{
				Line 0 "Domain Controllers: "
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Domain Controllers"
			}

			$DomainControllers = $Null
			$DomainControllers = Get-ADDomainController -Filter * -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Name
			
			If($? -and $Null -ne $DomainControllers)
			{
				$DomainControllers | ForEach-Object{$Script:AllDomainControllers.Add($_) | Out-Null} #3.05 fix by Jorge de Almeida Pinto
				#$Script:AllDomainControllers.Add($DomainControllers) > $Null # <= JORGE: HERE THE COLLETION OF OBJECTS IS ADDED AS ONE OBJECT!
				#$Script:AllDomainControllers = $Script:AllDomainControllers | Sort-Object Name -Unique #remove duplicates now that this can be done three times

				If($MSWord -or $PDF)
				{
					[System.Collections.Hashtable[]] $WordTable = @();
					ForEach($DomainController in $DomainControllers)
					{
						$WordTableRowHash = @{
						DCName = $DomainController.Name; 
						}
						$WordTable += $WordTableRowHash;
					}
					
					#set column widths
					If($WordTable.Count -gt 0)	#3.07
					{
						$Table = AddWordTable -Hashtable $WordTable `
						-Columns  DCName `
						-Headers "Name" `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 105;
						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
				}
				If($Text)
				{
					ForEach($DomainController in $DomainControllers)
					{
						Line 1 $DomainController.Name
					}
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					ForEach($DomainController in $DomainControllers)
					{
						$rowdata += @(,($DomainController.Name,$htmlwhite))
					}

					$columnHeaders = @("Name",$htmlsb)
					$columnWidths = @("105")
					FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "105"
					WriteHTMLLine 0 0 ''

					$rowdata = $null
				}
			}
			ElseIf(!$?)
			{
				Write-Warning "Error retrieving domain controller data for domain $($Domain)"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Error retrieving domain controller data for domain $($Domain)" "" $Null 0 $False $True
				}
				If($Text)
				{
					Line 0 "Error retrieving domain controller data for domain $($Domain)"
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Error retrieving domain controller data for domain $($Domain)" -options $htmlBold
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No Domain controller data was retrieved for domain $($Domain)" "" $Null 0 $False $True
				}
				If($Text)
				{
					Line 0 "No Domain controller data was retrieved for domain $($Domain)"
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No Domain controller data was retrieved for domain $($Domain)" -options $htmlBold
				}
			}

			Write-Verbose "$(Get-Date -Format G): `t`tProcessing Fine Grained Password Policies"
			
			#are FGPP cmdlets available
			If(Get-Command -Name "Get-ADFineGrainedPasswordPolicy" -ea 0)
			{
				#3.11 move the headings out so the error or notice hss a section heading
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "Fine Grained Password Policies"
				}
				If($Text)
				{
					Line 0 "Fine Grained Password Policies"
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Fine Grained Password Policies"
				}
				
				$FGPPs = $Null
				$FGPPs = Get-ADFineGrainedPasswordPolicy -Searchbase $DomainInfo.DistinguishedName -Filter * -Properties * -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Precedence, ObjectGUID
				
				If($? -and $Null -ne $FGPPs)
				{
					ForEach($FGPP in $FGPPs)
					{
						If($MSWord -or $PDF)
						{
							[System.Collections.Hashtable[]] $ScriptInformation = @()
							$ScriptInformation += @{ Data = "Name"; Value = $FGPP.Name; }
							$ScriptInformation += @{ Data = "Precedence"; Value = $FGPP.Precedence.ToString(); }
							
							If($FGPP.MinPasswordLength -eq 0)
							{
								$ScriptInformation += @{ Data = "Enforce minimum password length"; Value = "Not enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Enforce minimum password length"; Value = "Enabled"; }
								$ScriptInformation += @{ Data = "     Minimum password length (characters)"; Value = $FGPP.MinPasswordLength.ToString(); }
							}
							
							If($FGPP.PasswordHistoryCount -eq 0)
							{
								$ScriptInformation += @{ Data = "Enforce password history"; Value = "Not enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Enforce password history"; Value = "Enabled"; }
								$ScriptInformation += @{ Data = "     Number of passwords remembered"; Value = $FGPP.PasswordHistoryCount.ToString(); }
							}
							
							If($FGPP.ComplexityEnabled -eq $True)
							{
								$ScriptInformation += @{ Data = "Password must meet complexity requirements"; Value = "Enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Password must meet complexity requirements"; Value = "Not enabled"; }
							}
							
							If($FGPP.ReversibleEncryptionEnabled -eq $True)
							{
								$ScriptInformation += @{ Data = "Store password using reversible encryption"; Value = "Enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Store password using reversible encryption"; Value = "Not enabled"; }
							}
							
							If($FGPP.ProtectedFromAccidentalDeletion -eq $True)
							{
								$ScriptInformation += @{ Data = "Protect from accidental deletion"; Value = "Enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Protect from accidental deletion"; Value = "Not enabled"; }
							}
							
							$ScriptInformation += @{ Data = "Password age options"; Value = ""; }
							If($FGPP.MinPasswordAge.Days -eq 0)
							{
								$ScriptInformation += @{ Data = "     Enforce minimum password age"; Value = "Not enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "     Enforce minimum password age"; Value = "Enabled"; }
								$ScriptInformation += @{ Data = "          User cannot change the password within (days)"; Value = $FGPP.MinPasswordAge.TotalDays.ToString(); }
							}
							
							If($FGPP.MaxPasswordAge -eq 0)
							{
								$ScriptInformation += @{ Data = "     Enforce maximum password age"; Value = "Not enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "     Enforce maximum password age"; Value = "Enabled"; }
								$ScriptInformation += @{ Data = "          User must change the password after (days)"; Value = $FGPP.MaxPasswordAge.TotalDays.ToString(); }
							}
							
							If($FGPP.LockoutThreshold -eq 0)
							{
								$ScriptInformation += @{ Data = "Enforce account lockout policy"; Value = "Not enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Enforce account lockout policy"; Value = "Enabled"; }
								$ScriptInformation += @{ Data = "     Number of failed logon attempts allowed"; Value = $FGPP.LockoutThreshold.ToString(); }
								$ScriptInformation += @{ Data = "     Reset failed logon attempts count after (mins)"; Value = $FGPP.LockoutObservationWindow.TotalMinutes.ToString(); }
								If($FGPP.LockoutDuration -eq 0)
								{
									$ScriptInformation += @{ Data = "     Account will be locked out"; Value = ""; }
									$ScriptInformation += @{ Data = "          Until an administrator manually unlocks the account"; Value = ""; }
								}
								Else
								{
									$ScriptInformation += @{ Data = "     Account will be locked out for a duration of (mins)"; Value = $FGPP.LockoutDuration.TotalMinutes.ToString(); }
								}
								
							}
							
							$ScriptInformation += @{ Data = "Description"; Value = $FGPP.Description; }
							
							#$results = Get-ADFineGrainedPasswordPolicySubject -Identity $FGPP.Name -EA 0 | Sort-Object Name
							$results = Get-ADFineGrainedPasswordPolicySubject -Identity $FGPP.Name -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Name #3.05 fix by Jorge de Almeida Pinto
							
							If($? -and $Null -ne $results)
							{
								$cnt = 0
								ForEach($Item in $results)
								{
									$cnt++
									
									If($cnt -eq 1)
									{
										$ScriptInformation += @{ Data = "Directly Applies To"; Value = $Item.Name; }
									}
									Else
									{
										$ScriptInformation += @{ Data = ""; Value = $($Item.Name); }
									}
								}
							}
							Else
							{
							}
							
							$Table = AddWordTable -Hashtable $ScriptInformation `
							-Columns Data,Value `
							-List `
							-Format $wdTableGrid `
							-AutoFit $wdAutoFitFixed;

							SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

							$Table.Columns.Item(1).Width = 275;
							$Table.Columns.Item(2).Width = 200;

							$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

							FindWordDocumentEnd
							$Table = $Null
							WriteWordLine 0 0 ""
						}
						If($Text)
						{
							Line 1 "Name`t`t`t`t`t`t`t`t: " $FGPP.Name
							Line 1 "Precedence`t`t`t`t`t`t`t: " $FGPP.Precedence.ToString()
							
							Line 1 "Enforce minimum password length`t`t`t`t`t: " -NoNewLine
							If($FGPP.MinPasswordLength -eq 0)
							{
								Line 0 "Not enabled"
							}
							Else
							{
								Line 0 "Enabled"
								Line 3 "Minimum password length (characters)`t`t: " $FGPP.MinPasswordLength.ToString()
							}
							
							Line 1 "Enforce password history`t`t`t`t`t: " -NoNewLine
							If($FGPP.PasswordHistoryCount -eq 0)
							{
								Line 0 "Not enabled"
							}
							Else
							{
								Line 0 "Enabled"
								Line 3 "Number of passwords remembered`t`t`t: " $FGPP.PasswordHistoryCount.ToString()
							}
							
							Line 1 "Password must meet complexity requirements`t`t`t: " -NoNewLine
							If($FGPP.ComplexityEnabled -eq $True)
							{
								Line 0 "Enabled"
							}
							Else
							{
								Line 0 "Not enabled"
							}
							
							Line 1 "Store password using reversible encryption`t`t`t: " -NoNewLine
							If($FGPP.ReversibleEncryptionEnabled -eq $True)
							{
								Line 0 "Enabled"
							}
							Else
							{
								Line 0 "Not enabled"
							}
							
							Line 1 "Protect from accidental deletion`t`t`t`t: " -NoNewLine
							If($FGPP.ProtectedFromAccidentalDeletion -eq $True)
							{
								Line 0 "Enabled"
							}
							Else
							{
								Line 0 "Not enabled"
							}
							
							Line 1 "Password age options:"
							If($FGPP.MinPasswordAge.Days -eq 0)
							{
								Line 2 "Enforce minimum password age`t`t`t`t: Not enabled"
							}
							Else
							{
								Line 2 "Enforce minimum password age`t`t`t`t: Enabled" 
								Line 3 "User cannot change the password within (days)`t: " $FGPP.MinPasswordAge.TotalDays.ToString()
							}
							
							If($FGPP.MaxPasswordAge -eq 0)
							{
								Line 2 "Enforce maximum password age`t`t`t`t: Not enabled"
							}
							Else
							{
								Line 2 "Enforce maximum password age`t`t`t`t: Enabled"
								Line 3 "User must change the password after (days)`t: " $FGPP.MaxPasswordAge.TotalDays.ToString()
							}
							
							Line 1 "Enforce account lockout policy`t`t`t`t`t: " -NoNewLine
							If($FGPP.LockoutThreshold -eq 0)
							{
								Line 0 "Not enabled"
							}
							Else
							{
								Line 0 "Enabled"
								Line 2 "Number of failed logon attempts allowed`t`t`t: " $FGPP.LockoutThreshold.ToString()
								Line 2 "Reset failed logon attempts count after (mins)`t`t: " $FGPP.LockoutObservationWindow.TotalMinutes.ToString()
								If($FGPP.LockoutDuration -eq 0)
								{
									Line 2 "Account will be locked out"
									Line 3 "Until an administrator manually unlocks the account"
								}
								Else
								{
									Line 2 "Account will be locked out for a duration of (mins)`t: " $FGPP.LockoutDuration.TotalMinutes.ToString()
								}
								
							}
							
							Line 1 "Description`t`t`t`t`t`t`t: " $FGPP.Description
							
							#$results = Get-ADFineGrainedPasswordPolicySubject -Identity $FGPP.Name -EA 0 | Sort-Object Name
							$results = Get-ADFineGrainedPasswordPolicySubject -Identity $FGPP.Name -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Name #3.05 fix by Jorge de Almeida Pinto
							
							If($? -and $Null -ne $results)
							{
								Line 1 "Directly Applies To`t`t`t`t`t`t: " -NoNewLine
								$cnt = 0
								ForEach($Item in $results)
								{
									$cnt++
									
									If($cnt -eq 1)
									{
										Line 0 $Item.Name
									}
									Else
									{
										Line 9 "  $($Item.Name)"
									}
								}
							}
							Else
							{
							}
							Line 0 ""
						}
						If($HTML)
						{
							$rowdata = @()
							$columnHeaders = @("Precedence",$htmlsb,$FGPP.Precedence.ToString(),$htmlwhite)
							
							If($FGPP.MinPasswordLength -eq 0)
							{
								$rowdata += @(,("Enforce minimum password length",$htmlsb,"Not enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Enforce minimum password length",$htmlsb,"Enabled",$htmlwhite))
								$rowdata += @(,("     Minimum password length (characters)",$htmlsb,$FGPP.MinPasswordLength.ToString(),$htmlwhite))
							}
							
							If($FGPP.PasswordHistoryCount -eq 0)
							{
								$rowdata += @(,("Enforce password history",$htmlsb,"Not enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Enforce password history",$htmlsb,"Enabled",$htmlwhite))
								$rowdata += @(,("     Number of passwords remembered",$htmlsb,$FGPP.PasswordHistoryCount.ToString(),$htmlwhite))
							}
							
							If($FGPP.ComplexityEnabled -eq $True)
							{
								$rowdata += @(,("Password must meet complexity requirements",$htmlsb,"Enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Password must meet complexity requirements",$htmlsb,"Not enabled",$htmlwhite))
							}
							
							If($FGPP.ReversibleEncryptionEnabled -eq $True)
							{
								$rowdata += @(,("Store password using reversible encryption",$htmlsb,"Enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Store password using reversible encryption",$htmlsb,"Not enabled",$htmlwhite))
							}
							
							If($FGPP.ProtectedFromAccidentalDeletion -eq $True)
							{
								$rowdata += @(,("Protect from accidental deletion",$htmlsb,"Enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Protect from accidental deletion",$htmlsb,"Not enabled",$htmlwhite))
							}
							
							$rowdata += @(,("Password age options",$htmlsb,"",$htmlwhite))
							If($FGPP.MinPasswordAge.Days -eq 0)
							{
								$rowdata += @(,("     Enforce minimum password age",$htmlsb,"Not enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("     Enforce minimum password age",$htmlsb,"Enabled",$htmlwhite))
								$rowdata += @(,("          User cannot change the password within (days)",$htmlsb,$FGPP.MinPasswordAge.TotalDays.ToString(),$htmlwhite))
							}
							
							If($FGPP.MaxPasswordAge -eq 0)
							{
								$rowdata += @(,("     Enforce maximum password age",$htmlsb,"Not enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("     Enforce maximum password age",$htmlsb,"Enabled",$htmlwhite))
								$rowdata += @(,("          User must change the password after (days)",$htmlsb,$FGPP.MaxPasswordAge.TotalDays.ToString(),$htmlwhite))
							}
							
							If($FGPP.LockoutThreshold -eq 0)
							{
								$rowdata += @(,("Enforce account lockout policy",$htmlsb,"Not enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Enforce account lockout policy",$htmlsb,"Enabled",$htmlwhite))
								$rowdata += @(,("     Number of failed logon attempts allowed",$htmlsb,$FGPP.LockoutThreshold.ToString(),$htmlwhite))
								$rowdata += @(,("     Reset failed logon attempts count after (mins)",$htmlsb,$FGPP.LockoutObservationWindow.TotalMinutes.ToString(),$htmlwhite))
								If($FGPP.LockoutDuration -eq 0)
								{
									$rowdata += @(,("     Account will be locked out",$htmlsb,"",$htmlwhite))
									$rowdata += @(,("          Until an administrator manually unlocks the account",$htmlsb,"",$htmlwhite))
								}
								Else
								{
									$rowdata += @(,("     Account will be locked out for a duration of (mins)",$htmlsb,$FGPP.LockoutDuration.TotalMinutes.ToString(),$htmlwhite))
								}
								
							}
							
							$rowdata += @(,("Description",$htmlsb,$FGPP.Description,$htmlwhite))
							
							#$results = Get-ADFineGrainedPasswordPolicySubject -Identity $FGPP.Name -EA 0 | Sort-Object Name
							$results = Get-ADFineGrainedPasswordPolicySubject -Identity $FGPP.Name -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Name #3.05 fix by Jorge de Almeida Pinto
							
							If($? -and $Null -ne $results)
							{
								$cnt = 0
								ForEach($Item in $results)
								{
									$cnt++
									
									If($cnt -eq 1)
									{
										$rowdata += @(,("Directly Applies To",$htmlsb,$Item.Name,$htmlwhite))
									}
									Else
									{
										$rowdata += @(,("",$htmlsb,$($Item.Name),$htmlwhite))
									}
								}
							}
							Else
							{
							}
							
							$msg = "Name: $($FGPP.Name)"
							$columnWidths = @("500","100")
							FormatHTMLTable -tableHeader $msg `
								-rowArray $rowdata `
								-columnArray $columnHeaders `
								-fixedWidth $columnWidths `
								-tablewidth "600"
							WriteHTMLLine 0 0 ''

							$rowData = $null
						}
					}
				}
				ElseIf(!$?)
				{
					Write-Warning "Error retrieving Fine Grained Password Policy data for domain $($Domain)"
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 0 "Error retrieving Fine Grained Password Policy data for domain $($Domain)"
						WriteWordLine 0 0 ""
					}
					If($Text)
					{
						Line 0 "Error retrieving Fine Grained Password Policy data for domain $($Domain)"
						Line 0 ""
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 "Error retrieving Fine Grained Password Policy data for domain $($Domain)"
						WriteHTMLLine 0 0 " "
					}
				}
				Else
				{
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 0 "No Fine Grained Password Policy data was retrieved for domain $($Domain)"
						WriteWordLine 0 0 ""
					}
					If($Text)
					{
						Line 0 "No Fine Grained Password Policy data was retrieved for domain $($Domain)"
						Line 0 ""
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 "No Fine Grained Password Policy data was retrieved for domain $($Domain)"
						WriteHTMLLine 0 0 " "
					}
				}
			}
			Else
			{
				#FGPP cmdlets are not available
			}

			$First = $False
		}
	}
	$ADDomainTrusts = $Null
	$ADSchemaInfo = $Null
	$ChildDomains = $Null
	$DNSSuffixes = $Null
	$DomainControllers = $Null
	$ExchangeSchemaInfo = $Null
	$FGPPs = $Null
	$First = $Null
	$ReadOnlyReplicas = $Null
	$Replicas = $Null
	$SubordinateReferences = $Null
	$Table = $Null
}
#endregion

#region domain controllers
Function ProcessDomainControllers
{
	Write-Verbose "$(Get-Date -Format G): Writing domain controller data"

	#If($MSWORD -or $PDF)
	#{
	#	$Script:selection.InsertNewPage()
	#	WriteWordLine 1 0 "Domain Controllers in $($Script:ForestName)"
	#}
	#If($Text)
	#{
	#	Line 0 "///  Domain Controllers in $($Script:ForestName)  \\\"
	#}
	#If($HTML)
	#{
	#	WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Controllers in $($Script:ForestName)&nbsp;&nbsp;\\\"
	#}

	#$Script:AllDomainControllers = $Script:AllDomainControllers | Sort-Object Name
	$Script:AllDomainControllers = $Script:AllDomainControllers | Sort-Object Domain,Name #changed in 3.04
	$First = $True
	$SaveDomain = "" #added in 3.04

	ForEach($DC in $Script:AllDomainControllers)
	{
		#3.04
		#process DCs by domain. Don't put every DC in the root domain
		$DomainName = $DC.Domain
		If($SaveDomain -ne $DomainName)
		{
			If($MSWORD -or $PDF)
			{
				$Script:selection.InsertNewPage()
				WriteWordLine 1 0 "Domain Controllers in $($DomainName)"
			}
			If($Text)
			{
				Line 0 "///  Domain Controllers in $($DomainName)  \\\"
			}
			If($HTML)
			{
				WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Controllers in $($DomainName)&nbsp;&nbsp;\\\"
			}
			$First = $True
		}
		Write-Verbose "$(Get-Date -Format G): `tProcessing domain controller $($DC.name)"
		$FSMORoles = $DC.OperationMasterRoles | Sort-Object 
		$Partitions = $DC.Partitions | Sort-Object 
		
		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each DC, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}
		
		If($MSWORD -or $PDF)
		{
			WriteWordLine 2 0 $DC.Name
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$WordHighlightedCells = New-Object System.Collections.ArrayList
			[int] $CurrentServiceIndex = 1
			$ScriptInformation += @{ Data = "Computer Object DN"; Value = $DC.ComputerObjectDN; }
			If($DC.ComputerObjectDN -NotLike "*OU=Domain Controllers*")
			{
				$WordHighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 2; }) > $Null
			}
			$ScriptInformation += @{ Data = "Default partition"; Value = $DC.DefaultPartition; }
			$ScriptInformation += @{ Data = "Domain"; Value = $DC.domain; }
			If($DC.Enabled -eq $True)
			{
				$tmp = "True"
			}
			Else
			{
				$tmp = "False"
			}
			$ScriptInformation += @{ Data = "Enabled"; Value = $tmp; }
			$ScriptInformation += @{ Data = "Hostname"; Value = $DC.HostName; }
			If($DC.IsGlobalCatalog -eq $True)
			{
				$tmp = "Yes" 
			}
			Else
			{
				$tmp = "No"
			}
			$ScriptInformation += @{ Data = "Global Catalog"; Value = $tmp; }
			If($DC.IsReadOnly -eq $True)
			{
				$tmp = "Yes"
			}
			Else
			{
				$tmp = "No"
			}
			$ScriptInformation += @{ Data = "Read-only"; Value = $tmp; }
			$ScriptInformation += @{ Data = "LDAP port"; Value = $DC.LdapPort.ToString(); }
			$ScriptInformation += @{ Data = "SSL port"; Value = $DC.SslPort.ToString(); }
			If($Null -eq $FSMORoles)
			{
				$ScriptInformation += @{ Data = "Operation Master roles"; Value = "<None>"; }
			}
			Else
			{
				$cnt = 0
				ForEach($FSMORole in $FSMORoles)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						$ScriptInformation += @{ Data = "Operation Master roles"; Value = $FSMORole.ToString(); }
					}
					Else
					{
						$ScriptInformation += @{ Data = ""; Value = $FSMORole.ToString(); }
					}
				}
			}
			If($Null -eq $Partitions)
			{
				$ScriptInformation += @{ Data = "Partitions"; Value = "<None>"; }
			}
			Else
			{
				$cnt = 0
				ForEach($Partition in $Partitions)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						$ScriptInformation += @{ Data = "Partitions"; Value = $Partition.ToString(); }
					}
					Else
					{
						$ScriptInformation += @{ Data = ""; Value = $Partition.ToString(); }
					}
					
				}
			}
			$ScriptInformation += @{ Data = "Site"; Value = $DC.Site; }
			$ScriptInformation += @{ Data = "Operating System"; Value = $DC.OperatingSystem; }
			
			If(![String]::IsNullOrEmpty($DC.OperatingSystemServicePack))
			{
				$ScriptInformation += @{ Data = "Service Pack"; Value = $DC.OperatingSystemServicePack; }
			}
			$ScriptInformation += @{ Data = "Operating System version"; Value = $DC.OperatingSystemVersion; }
			
			If(!$Hardware)
			{
				If([String]::IsNullOrEmpty($DC.IPv4Address))
				{
					$tmp = "<None>"
				}
				Else
				{
					$tmp = $DC.IPv4Address
				}
				$ScriptInformation += @{ Data = "IPv4 Address"; Value = $tmp; }

				If([String]::IsNullOrEmpty($DC.IPv6Address))
				{
					$tmp = "<None>"
				}
				Else
				{
					$tmp = $DC.IPv6Address
				}
				$ScriptInformation += @{ Data = "IPv6 Address"; Value = $tmp; }
			}
			
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
			If($WordHighlightedCells.Count -gt 0)
			{
				SetWordCellFormat -Coordinates $WordHighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;
			}

			$Table.Columns.Item(1).Width = 140;
			$Table.Columns.Item(2).Width = 300;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 "///  DC: $($DC.Name)  \\\"
			If($DC.ComputerObjectDN -NotLike "*OU=Domain Controllers*")
			{
				Line 1 "Computer Object DN`t`t: ***$($DC.ComputerObjectDN)***"
			}
			Else
			{
				Line 1 "Computer Object DN`t`t: " $DC.ComputerObjectDN
			}
			Line 1 "Default partition`t`t: " $DC.DefaultPartition
			Line 1 "Domain`t`t`t`t: " $DC.domain
			If($DC.Enabled -eq $True)
			{
				$tmp = "True"
			}
			Else
			{
				$tmp = "False"
			}
			Line 1 "Enabled`t`t`t`t: " $tmp
			Line 1 "Hostname`t`t`t: " $DC.HostName
			If($DC.IsGlobalCatalog -eq $True)
			{
				$tmp = "Yes" 
			}
			Else
			{
				$tmp = "No"
			}
			Line 1 "Global Catalog`t`t`t: " $tmp
			If($DC.IsReadOnly -eq $True)
			{
				$tmp = "Yes"
			}
			Else
			{
				$tmp = "No"
			}
			Line 1 "Read-only`t`t`t: " $tmp
			Line 1 "LDAP port`t`t`t: " $DC.LdapPort.ToString()
			Line 1 "SSL port`t`t`t: " $DC.SslPort.ToString()
			Line 1 "Operation Master roles`t`t: " -NoNewLine
			If($Null -eq $FSMORoles)
			{
				Line 0 "<None>"
			}
			Else
			{
				$cnt = 0
				ForEach($FSMORole in $FSMORoles)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						Line 0 $FSMORole.ToString()
					}
					Else
					{
						Line 5 "  $($FSMORole.ToString())"
					}
					
				}
			}
			Line 1 "Partitions`t`t`t: " -NoNewLine
			If($Null -eq $Partitions)
			{
				Line 0 "<None>"
			}
			Else
			{
				$cnt = 0
				ForEach($Partition in $Partitions)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						Line 0 $Partition.ToString()
					}
					Else
					{
						Line 5 "  $($Partition.ToString())"
					}
				}
			}
			Line 1 "Site`t`t`t`t: " $DC.Site
			Line 1 "Operating System`t`t: " $DC.OperatingSystem
			If(![String]::IsNullOrEmpty($DC.OperatingSystemServicePack))
			{
				Line 1 "Service Pack`t`t`t: " $DC.OperatingSystemServicePack
			}
			Line 1 "Operating System version`t: " $DC.OperatingSystemVersion
			
			If(!$Hardware)
			{
				Line 1 "IPv4 Address`t`t`t: " -NoNewLine
				If([String]::IsNullOrEmpty($DC.IPv4Address))
				{
					Line 0 "<None>"
				}
				Else
				{
					Line 0 $DC.IPv4Address
				}
				Line 1 "IPv6 Address`t`t`t: " -NoNewLine
				If([String]::IsNullOrEmpty($DC.IPv6Address))
				{
					Line 0 "<None>"
				}
				Else
				{
					Line 0 $DC.IPv6Address
				}
			}
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($DC.Name)&nbsp;&nbsp;\\\"
			$rowdata = @()
			If($DC.ComputerObjectDN -NotLike "*OU=Domain Controllers*")
			{
				$HTMLHighlightedCells = $htmlred
			}
			Else
			{
				$HTMLHighlightedCells = $htmlwhite
			} 
			$columnHeaders = @("Computer Object DN",$htmlsb,$DC.DefaultPartition,$HTMLHighlightedCells)
			$rowdata += @(,("Default partition",$htmlsb,$DC.DefaultPartition,$htmlwhite))
			$rowdata += @(,('Domain',$htmlsb,$DC.domain,$htmlwhite))
			If($DC.Enabled -eq $True)
			{
				$tmp = "True"
			}
			Else
			{
				$tmp = "False"
			}
			$rowdata += @(,('Enabled',$htmlsb,$tmp,$htmlwhite))
			$rowdata += @(,('Hostname',$htmlsb,$DC.HostName,$htmlwhite))
			If($DC.IsGlobalCatalog -eq $True)
			{
				$tmp = "Yes" 
			}
			Else
			{
				$tmp = "No"
			}
			$rowdata += @(,('Global Catalog',$htmlsb,$tmp,$htmlwhite))
			If($DC.IsReadOnly -eq $True)
			{
				$tmp = "Yes"
			}
			Else
			{
				$tmp = "No"
			}
			$rowdata += @(,('Read-only',$htmlsb,$tmp,$htmlwhite))
			$rowdata += @(,('LDAP port',$htmlsb,$DC.LdapPort.ToString(),$htmlwhite))
			$rowdata += @(,('SSL port',$htmlsb,$DC.SslPort.ToString(),$htmlwhite))
			If($Null -eq $FSMORoles)
			{
				$rowdata += @(,('Operation Master roles',$htmlsb,"None",$htmlwhite))
			}
			Else
			{
				$cnt = 0
				ForEach($FSMORole in $FSMORoles)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						$rowdata += @(,('Operation Master roles',$htmlsb,$FSMORole.ToString(),$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('',$htmlsb,$FSMORole.ToString(),$htmlwhite))
					}
				}
			}
			If($Null -eq $Partitions)
			{
				$rowdata += @(,('Partitions',$htmlsb,"None",$htmlwhite))
			}
			Else
			{
				$cnt = 0
				ForEach($Partition in $Partitions)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						$rowdata += @(,('Partitions',$htmlsb,$Partition.ToString(),$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('',$htmlsb,$Partition.ToString(),$htmlwhite))
					}
				}
			}
			$rowdata += @(,('Site',$htmlsb,$DC.Site,$htmlwhite))
			$rowdata += @(,('Operating System',$htmlsb,$DC.OperatingSystem,$htmlwhite))
			
			If(![String]::IsNullOrEmpty($DC.OperatingSystemServicePack))
			{
				$rowdata += @(,('Service Pack',$htmlsb,$DC.OperatingSystemServicePack,$htmlwhite))
			}
			$rowdata += @(,('Operating System version',$htmlsb,$DC.OperatingSystemVersion,$htmlwhite))
			
			If(!$Hardware)
			{
				If([String]::IsNullOrEmpty($DC.IPv4Address))
				{
					$tmp = "None"
				}
				Else
				{
					$tmp = $DC.IPv4Address
				}
				$rowdata += @(,('IPv4 Address',$htmlsb,$tmp,$htmlwhite))

				If([String]::IsNullOrEmpty($DC.IPv6Address))
				{
					$tmp = "None"
				}
				Else
				{
					$tmp = $DC.IPv6Address
				}
				$rowdata += @(,('IPv6 Address',$htmlsb,$tmp,$htmlwhite))
			}
			
			$columnWidths = @("175","300")
			FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "475"
			#WriteHTMLLine 0 0 ' '

			$rowData = $null
		}
		
		If($Script:DARights -and $Script:Elevated)
		{
			#3.06, add testing to see if the DC is online
			If(Test-NetConnection -ComputerName $DC.HostName -Port 88 -InformationLevel Quiet -EA 0) #port 88 is the KDC and is unique to DCs (thanks to Matthew Woolnough)
			{
				OutputTimeServerRegistryKeys $DC.HostName
		
				OutputADFileLocations $DC.HostName
			
				OutputEventLogInfo $DC.HostName
			}
			Else
			{
				$txt = "$(Get-Date): `t`t$($DC.Name) is offline or unreachable.  Time Server, AD File Locations, and Event Log data is skipped."
				Write-Verbose $txt
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt
				}
				If($Text)
				{
					Line 0 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}
			}
		}
		
		If($Hardware -or $Services -or $DCDNSInfo)
		{
			#If(Test-Connection -ComputerName $DC.HostName -quiet -EA 0)
			#3.06, change from $ComputerName to $DC.HostName
			If(Test-NetConnection -ComputerName $DC.HostName -Port 88 -InformationLevel Quiet -EA 0) #port 88 is the KDC and is unique to DCs (thanks to Matthew Woolnough)
			{
				If($Hardware)
				{
					GetComputerWMIInfo $DC.HostName
				}
				
				If($DCDNSInfo)
				{
					BuildDCDNSIPConfigTable $DC.HostName $DC.Site
				}

				If($Services)
				{
					GetComputerServices $DC.HostName
				}
			}
			Else
			{
				$txt = "$(Get-Date): `t`t$($DC.Name) is offline or unreachable.  Hardware inventory is skipped."
				Write-Verbose $txt
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt
				}
				If($Text)
				{
					Line 0 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}

				If($Hardware -and -not $Services)
				{
					$txt = "Hardware inventory was skipped."
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 0 $txt
					}
					If($Text)
					{
						Line 0 $txt
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 $txt
					}
				}
				ElseIf($Services -and -not $Hardware)
				{
					$txt = "Services was skipped."
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 0 $txt
					}
					If($Text)
					{
						Line 0 $txt
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 $txt
					}
				}
				ElseIf($Hardware -and $Services)
				{
					$txt = "Hardware inventory and Services were skipped."
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 0 $txt
					}
					If($Text)
					{
						Line 0 $txt
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 $txt
					}
				}
			}
		}
		$First = $False
		$SaveDomain = $DomainName
	}
	$Script:AllDomainControllers = $Null
} ## end Function ProcessDomainControllers

Function OutputTimeServerRegistryKeys 
{
	Param
	(
		[String] $DCName
	)
	
	Write-Verbose "$(Get-Date -Format G): `t`tTimeServer Registry Keys"
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config	AnnounceFlags
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config	MaxNegPhaseCorrection
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config	MaxPosPhaseCorrection
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Parameters	NtpServer
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Parameters	Type 	
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient	SpecialPollInterval
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\VMICTimeProvider Enabled
	
	$AnnounceFlags = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" "AnnounceFlags" $DCName
	#If( $null -eq $AnnounceFlags -and $error.Count -gt 0 -and $error[ 0 ].Exception.HResult -eq -2146233087 )
	#3.06 check only for null to catch appliances (Riverbed) with no registry
	If( $null -eq $AnnounceFlags )
	{
		## DCName can't be contacted or DCName is an appliance with no registry
		$AnnounceFlags = 'n/a'
		$MaxNegPhaseCorrection = 'n/a'
		$MaxPosPhaseCorrection = 'n/a'
		$NtpServer = 'n/a'
		$NtpType = 'n/a'
		$SpecialPollInterval = 'n/a'
		$VMICTimeProviderEnabled = 'n/a'
		$NTPSource = 'Cannot retrieve data from registry'
	}
	Else
	{
		$MaxNegPhaseCorrection = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" "MaxNegPhaseCorrection" $DCName
		$MaxPosPhaseCorrection = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" "MaxPosPhaseCorrection" $DCName
		$NtpServer = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" "NtpServer" $DCName
		$NtpType = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" "Type" $DCName
		$SpecialPollInterval = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient" "SpecialPollInterval" $DCName
		$VMICTimeProviderEnabled = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\VMICTimeProvider" "Enabled" $DCName
		$NTPSource = Invoke-Command -ComputerName $DCName {w32tm /query /computer:$DCName /source}
	}

	If( $VMICTimeProviderEnabled -eq 'n/a' )
	{
		$VMICEnabled = 'n/a'
	}
	ElseIf( $VMICTimeProviderEnabled -eq 0 )
	{
		$VMICEnabled = 'Disabled'
	}
	Else
	{
		$VMICEnabled = 'Enabled'
	}
	
	## create time server info array
	## after testing - it appears that a normal hashtable is much faster
	## than a PSCustomObject - 2019/03/10 - MBS
	## but a regular hashtable doesn't sort, a PSCO does - 2019/03/12 - MBS
	$obj = [PSCustomObject] @{
		DCName                = $DCName
		TimeSource            = $NTPSource
		AnnounceFlags         = $AnnounceFlags
		MaxNegPhaseCorrection = $MaxNegPhaseCorrection
		MaxPosPhaseCorrection = $MaxPosPhaseCorrection
		NtpServer             = $NtpServer
		NtpType               = $NtpType
		SpecialPollInterval   = $SpecialPollInterval
		VMICTimeProvider      = $VMICEnabled
	}

	$null = $Script:TimeServerInfo.Add( $obj )
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 "Time Server Information"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		
		$Scriptinformation += @{ Data = "Time source"; Value = $NTPSource; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\Config\AnnounceFlags"; Value = $AnnounceFlags; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\Config\MaxNegPhaseCorrection"; Value = $MaxNegPhaseCorrection; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\Config\MaxPosPhaseCorrection"; Value = $MaxPosPhaseCorrection; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\Parameters\NtpServer"; Value = $NtpServer; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\Parameters\Type"; Value = $NtpType; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\TimeProviders\NtpClient\SpecialPollInterval"; Value = $SpecialPollInterval; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\TimeProviders\VMICTimeProvider\Enabled"; Value = $VMICEnabled; }
		
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed

		SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

		$Table.Columns.Item(1).Width = 350
		$Table.Columns.Item(2).Width = 130

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Time Server Information"
		Line 0 ""
		Line 1 "Time source: " $NTPSource
		Line 1 "HKLM:\SYSTEM\CCS\Services\W32Time\"
		Line 2 "Config\AnnounceFlags`t`t`t`t: " $AnnounceFlags
		Line 2 "Config\MaxNegPhaseCorrection`t`t`t: " $MaxNegPhaseCorrection
		Line 2 "Config\MaxPosPhaseCorrection`t`t`t: " $MaxPosPhaseCorrection
		Line 2 "Parameters\NtpServer`t`t`t`t: " $NtpServer
		Line 2 "Parameters\Type`t`t`t`t`t: " $NtpType
		Line 2 "TimeProviders\NtpClient\SpecialPollInterval`t: " $SpecialPollInterval
		Line 2 "TimeProviders\VMICTimeProvider\Enabled`t`t: " $VMICEnabled
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 'Time Server Information'
		#V3.00 pre-allocate rowdata
		$rowdata = New-Object System.Array[] 7

		$rowdata[ 0 ] = @(
			'HKLM:\SYSTEM\CCS\Services\W32Time\Config\AnnounceFlags', $htmlsb,
			$AnnounceFlags, $htmlwhite
		)

		$rowdata[ 1 ] = @(
			'HKLM:\SYSTEM\CCS\Services\W32Time\Config\MaxNegPhaseCorrection', $htmlsb,
			$MaxNegPhaseCorrection, $htmlwhite
		)

		$rowdata[ 2 ] = @(
			'HKLM:\SYSTEM\CCS\Services\W32Time\Config\MaxPosPhaseCorrection', $htmlsb,
			$MaxPosPhaseCorrection, $htmlwhite
		)

		$rowdata[ 3 ] = @(
			'HKLM:\SYSTEM\CCS\Services\W32Time\Parameters\NtpServer', $htmlsb,
			$NtpServer, $htmlwhite
		)

		$rowdata[ 4 ] = @(
			'HKLM:\SYSTEM\CCS\Services\W32Time\Parameters\Type', $htmlsb,
			$NtpType, $htmlwhite
		)

		$rowdata[ 5 ] = @(
			'HKLM:\SYSTEM\CCS\Services\W32Time\TimeProviders\NtpClient\SpecialPollInterval', $htmlsb,
			$SpecialPollInterval, $htmlwhite
		)

		$rowdata[ 6 ] = @(
			'HKLM:\SYSTEM\CCS\Services\W32Time\TimeProviders\VMICTimeProvider\Enabled', $htmlsb,
			$VMICEnabled, $htmlwhite
		)

		$columnWidths  = @( '350px', '300px' )
		$columnHeaders = @(
			'Time source', $htmlsb,
			$NTPSource,    $htmlwhite
		)

		FormatHTMLTable -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth '650'
		WriteHTMLLine 0 0 ''

		$rowData = $null
	}
} ## end Function OutputTimeServerRegistryKeys

Function OutputADFileLocations
{
	Param
	(
		[String] $DCName 
	)
	
	Write-Verbose "$(Get-Date -Format G): `t`tAD Database, Logfile and SYSVOL locations"
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NTDS\Parameters	'DSA Database file'
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NTDS\Parameters	'Database log files path'
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters	SysVol
	
	$DSADatabaseFile = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters" "DSA Database file" $DCName
	#If( $null -eq $DSADatabaseFile -and $error.Count -gt 0 -and $error[ 0 ].Exception.HResult -eq -2146233087 )
	#3.06 check only for null to catch appliances (Riverbed) with no registry
	If( $null -eq $DSADatabaseFile )
	{
		$DSADatabaseFile = ''
		$DatabaseLogFilesPath = ''
		$Sysvol = ''
		$SysvolState = 'Unknown'
	}
	Else
	{
		$DatabaseLogFilesPath = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters" "Database log files path" $DCName
		$Sysvol = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters" "SysVol" $DCName

		#$Result = Get-WMIObject -ComputerName $DCName -Namespace "root/microsoftdfs" -Class "dfsrreplicatedfolderinfo" -Filter "ReplicatedFolderName = 'SYSVOL Share'" -EA 0 | Select-Object State
		If($DCName -eq $env:computername)
		{
			$Result = Get-CIMInstance -Namespace "root/microsoftdfs" -Class "dfsrreplicatedfolderinfo" -Filter "ReplicatedFolderName = 'SYSVOL Share'" -EA 0 -Verbose:$False | Select-Object State
		}
		Else
		{
			$Result = Get-CIMInstance -CIMSession $DCName -Namespace "root/microsoftdfs" -Class "dfsrreplicatedfolderinfo" -Filter "ReplicatedFolderName = 'SYSVOL Share'" -EA 0 -Verbose:$False | Select-Object State
		}

		#original code before 3.07 change
		<#
		If($? -and $Null -ne $Result)
		{
			$SysvolState  = $Result.State

			#https://docs.microsoft.com/en-us/troubleshoot/windows-server/networking/troubleshoot-missing-sysvol-and-netlogon-shares
			#The state values can be any of:
			#0 = Uninitialized
			#1 = Initialized
			#2 = Initial Sync
			#3 = Auto Recovery
			#4 = Normal
			#5 = In Error

			Switch($Result.State)
			{
				0		{$SysvolState  = "0 = Uninitialized"; Break}
				1		{$SysvolState  = "1 = Initialized"; Break}
				2		{$SysvolState  = "2 = Initial Sync"; Break}
				3		{$SysvolState  = "3 = Auto Recovery"; Break}
				4		{$SysvolState  = "4 = Normal"; Break}
				5		{$SysvolState  = "5 = In Error"; Break}
				Default {$SysvolState  = "Unable to determine the SYSVOL State: $($Result.State)"; Break}
			}
		}
		Else
		{
			#$SysvolState  = "Unable to determine the SYSVOL State: $($Result.State)"
			$SysvolState  = "Unable to determine the SYSVOL State" #changed in 3.06, if $Result is null or an error happened, there is no State property
		}
		#>
		#end of original code before change
		
		If($?)
		{
			If($Null -ne $Result)
			{
				$SysvolState  = $Result.State

				#https://docs.microsoft.com/en-us/troubleshoot/windows-server/networking/troubleshoot-missing-sysvol-and-netlogon-shares
				#The state values can be any of:
				#0 = Uninitialized
				#1 = Initialized
				#2 = Initial Sync
				#3 = Auto Recovery
				#4 = Normal
				#5 = In Error

				Switch($Result.State)
				{
					0		{$SysvolState  = "0 = Uninitialized"; Break}
					1		{$SysvolState  = "1 = Initialized"; Break}
					2		{$SysvolState  = "2 = Initial Sync"; Break}
					3		{$SysvolState  = "3 = Auto Recovery"; Break}
					4		{$SysvolState  = "4 = Normal"; Break}
					5		{$SysvolState  = "5 = In Error"; Break}
					Default {$SysvolState  = "Unable to determine the SYSVOL State: $($Result.State)"; Break}
				}
			}
			Else
			{
				#Code supplied by Michael B. Smith
				If( Get-Command Get-SmbShare )
				{
					$Result = Get-SmbShare SYSVOL
					If( $? -and $null -ne $Result )
					{	
						$SysVolState = $Result.ShareState
					}
					Else
					{
						$SysvolState  = "Unable to determine the SYSVOL State (SMB)"
					}
				}
				Else
				{
					$Result = net.exe SHARE SYSVOL
					If( 0 -eq $LASTEXITCODE )
					{
						$SysVolState = 'Online'
					}
					ElseIf ( 2 -eq $LASTEXITCODE )
					{
						$SysVolState = 'No such share'
					}
					Else
					{
						$SysVolState ="Unable to determine the SYSVOL State (net.exe) $LASTEXITCODE"
					}
				}
				#end of MBS code
			}
		}
		Else
		{
			$SysvolState  = "Unable to determine the SYSVOL State" #changed in 3.06, if $Result is null or an error happened, there is no State property
		}
	}

	$obj = [PSCustomObject] @{
		DCName = $DCName
		SYSVOLState = $SysvolState
	}

	$null = $Script:DCSYSVOLState.Add( $obj )

	If( $DSADatabaseFile -and $DSADatabaseFile.Length -gt 0 )
	{
		$DITRemotePath = $DSADatabaseFile.Replace(":", "$")
		$DITFile = "\\$DCName\$DITRemotePath"
		$DITsize = ([System.IO.FileInfo]$DITFile).Length
		$DITsize = ($DITsize/1GB)
		$DSADatabaseFileSize = "{0:N3}" -f $DITsize
	}
	Else
	{
		$DITRemotePath = ''
		$DITFile       = ''
		$DITsize       = 0
		$DSADatabaseFileSize = '0.00'
	}

	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 "AD Database, Logfile and SYSVOL Locations"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$WordHighlightedCells = New-Object System.Collections.ArrayList
		[int] $CurrentServiceIndex = 5
		
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters\DSA Database file"; Value = $DSADatabaseFile; }
		$Scriptinformation += @{ Data = "DSA Database file size "; Value = "$($DSADatabaseFileSize) GB"; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters\Database log files path"; Value = $DatabaseLogFilesPath; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters\SysVol"; Value = $Sysvol; }
		$Scriptinformation += @{ Data = "SYSVOL State"; Value = $SysvolState; }
		If($SysvolState -NotLike "4 = Normal")
		{
			$WordHighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 2; }) > $Null
		}
		
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed

		SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15
		If($WordHighlightedCells.Count -gt 0)
		{
			SetWordCellFormat -Coordinates $WordHighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;
		}

		$Table.Columns.Item(1).Width = 350
		$Table.Columns.Item(2).Width = 130

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "AD Database, Logfile and SYSVOL Locations"
		Line 0 ""
		Line 1 "HKLM:\SYSTEM\CCS\Services\"
		Line 2 "NTDS\Parameters\DSA Database file`t: " $DSADatabaseFile
		Line 2 "DSA Database file size`t`t`t: $($DSADatabaseFileSize) GB"
		Line 2 "NTDS\Parameters\Database log files path`t: " $DatabaseLogFilesPath
		Line 2 "Netlogon\Parameters\SysVol`t`t: " $SysVol
		If($SysvolState -NotLike "4 = Normal")
		{
			Line 2 "SYSVOL State`t`t`t`t: " "***$($SysvolState)***"
		}
		Else
		{
			Line 2 "SYSVOL State`t`t`t`t: " $SysvolState
		}
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 'AD Database, Logfile and SYSVOL Locations'

		#V3.00 pre-allocate rowdata
		$rowdata = New-Object System.Array[] 4

		$columnHeaders = @('HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters\DSA Database file',$htmlsb,$DSADatabaseFile,$htmlwhite)
		$rowdata[ 0 ] = @('DSA Database file size',$htmlsb,"$DSADatabaseFileSize GB",$htmlwhite)
		$rowdata[ 1 ] = @('HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters\Database log files path',$htmlsb,$DatabaseLogFilesPath,$htmlwhite)
		$rowdata[ 2 ] = @('HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters\SysVol',$htmlsb,$SysVol,$htmlwhite)
		If($SysvolState -NotLike "4 = Normal")
		{
			$HTMLHighlightedCells = $htmlred
		}
		Else
		{
			$HTMLHighlightedCells = $htmlwhite
		} 
		$rowdata[ 3 ] = @('SYSVOL State',$htmlsb,$SysvolState,$HTMLHighlightedCells)

		$columnWidths  = @( '500', '150' )

		FormatHTMLTable -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth '650'
		WriteHTMLLine 0 0 ''

		$rowData = $null
	}
} ## end Function OutputADFileLocations

Function OutputEventLogInfo
{
	Param
	(
		[String] $DCName 
	)
	
	Write-Verbose "$(Get-Date -Format G): `t`tEvent Log Information"
	$ELInfo = $null ## New-Object System.Collections.ArrayList ## FIXME - make this an array instead of arraylist
	
	#V3.00 - note that we are sorted here by name, don't need to sort again later.
	#3.06 add try/catch
	
	try
	{
		$EventLogs = Get-EventLog -List -ComputerName $DCName -EA 0 | Select-Object MaximumKilobytes, Log | Sort-Object Log 
	}
	
	catch
	{
		$EventLogs = $Null
	}
	
	If($? -and $Null -ne $EventLogs)
	{
		$ELInfo = New-Object System.Array[] $EventLogs.Count
		$ELInx  = 0

		ForEach($EventLog in $EventLogs)
		{
			[String] $ELSize = "{0,10:N0}" -f $EventLog.MaximumKilobytes

			$obj = [PSCustomObject] @{
				DCName = $DCName
				EventLogName = $EventLog.Log
				EventLogSize = $ELSize
			}

			$null = $Script:DCEventLogInfo.Add( $obj )
			$ELInfo[ $ELInx ] = $obj
			$ELInx++
		}
	}
	Else
	{
		$ELInfo = New-Object System.Array[] 1

		[String] $ELSize = "{0,10:N0}" -f 0
	
		$obj = [PSCustomObject] @{
			DCName = $DCName
			EventLogName = 'Cannot retrieve event log information'
			EventLogSize = $ELSize
		}
		
		$null = $Script:DCEventLogInfo.Add( $obj )
		$ELInfo[ 0 ] = $obj
	}

	##v3.00 - doesn't need to be re-sorted or re-created
	##$xEventLogInfo = @($ELInfo | Sort-Object EventLogName)
	$xEventLogInfo = $ELInfo

	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Event Log Information"
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = $xEventLogInfo.Count + 1
		[int]$xRow = 1
		
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.AutoFitBehavior($wdAutoFitFixed)
		$Table.Style = $Script:MyHash.Word_TableGrid
	
		$Table.rows.first.headingformat = $wdHeadingFormatTrue
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

		$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Event Log Name"
		
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Text = "Event Log Size (KB)"
	}
	If($Text)
	{
		Line 0 "Event Log Information"
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Event Log Information"
		#V3.00 - pre-allocate rowdata
		## $rowdata = @()
		$rowData = New-Object System.Array[] $xEventLogInfo.Count
		$rowIndx = 0
	}

	ForEach($Item in $xEventLogInfo)
	{
		If($MSWord -or $PDF)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $Item.EventLogName
			$Table.Cell($xRow,2).Range.ParagraphFormat.Alignment = $wdCellAlignVerticalTop
			$Table.Cell($xRow,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
			$Table.Cell($xRow,2).Range.Text = $Item.EventLogSize
		}
		If($Text)
		{
			Line 1 "Event Log Name`t`t: " $Item.EventLogName
			Line 1 "Event Log Size (KB)`t: " $Item.EventLogSize
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata[ $rowIndx ] = @(
				$Item.EventLogName, $htmlwhite,
				$Item.EventLogSize, ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
			)
			$rowIndx++
		}
	}

	If($MSWord -or $PDF)
	{
		#set column widths
		$xcols = $table.columns

		ForEach($xcol in $xcols)
		{
			Switch ($xcol.Index)
			{
			  1 {$xcol.width = 150; Break}
			  2 {$xcol.width = 100; Break}
			}
		}
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitFixed)

		#Return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		#nothing to do
	}
	If($HTML)
	{
		$columnHeaders = @(
			'Event Log Name',      $htmlsb,
			'Event Log Size (KB)', $htmlsb
		)

		$columnWidths = @( '225px', '100px' )

		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '325'
		WriteHTMLLine 0 0 ''

		$rowData = $null
	}
} ## end Function OutputEventLogInfo
#endregion

#region organizational units
Function ProcessOrganizationalUnits
{
	Write-Verbose "$(Get-Date -Format G): Writing OU data by Domain"
	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Organizational Units"
	}
	If($Text)
	{
		Line 0 "///  Organizational Units  \\\"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Organizational Units&nbsp;&nbsp;\\\"
	}
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing domain $($Domain)"
		
		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			## FIXME-MBS: if Get-ADOrganizationalUnit fails, we could have
			## multiple blank pages in a row
			$Script:selection.InsertNewPage()
		}
		
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt = "OUs in Domain $($Domain) (Forest Root)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
			}
			If($Text)
			{
				Line 0 "///  $($txt)  \\\"
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
			}
		}
		Else
		{
			$txt = "OUs in Domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
			}
			If($Text)
			{
				Line 0 "///  $($txt)  \\\"
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
			}
		}
		
		#get all OUs for the domain
		#V3.00 - see optimizations applied to getDSUsers
		$OUs = @(Get-ADOrganizationalUnit -Filter * -Server $Domain `
		-Properties CanonicalName, DistinguishedName, Name, Created, ProtectedFromAccidentalDeletion -EA 0 | `
		Select-Object CanonicalName, DistinguishedName, Name, Created, ProtectedFromAccidentalDeletion | `
		Sort-Object CanonicalName)
		
		#V3.00 - simplify-logic - FIXME
		If( !$? )
		{
			$txt = "Error retrieving OU data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If( $null -eq $OUs )
		{
			$txt = "No OU data was retrieved for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		[int]$OUCount = 0
		[int]$NumOUs = $OUs.Count
		[int]$UnprotectedOUs = 0 #added in V2.22

		If($MSWORD -or $PDF)
		{
			$ItemsWordTable = New-Object System.Collections.ArrayList
			$WordHighlightedCells = New-Object System.Collections.ArrayList
			[int] $CurrentServiceIndex = 2
		}
		If($Text)
		{
			[int]$MaxOUNameLength = ($OUs.CanonicalName.SubString($OUs[0].CanonicalName.IndexOf("/")+1) | measure-object -maximum -property length).maximum
			
			If($MaxOUNameLength -gt 4) #4 is length of "Name"
			{
				#2 is to allow for spacing between columns
				Line 1 ("Name" + (' ' * ($MaxOUNameLength - 2))) -NoNewLine
				Line 0 "Created                Protected # Users # Computers # Groups"
				Line 1 ('=' * $MaxOUNameLength) -NoNewLine
				Line 0 "==============================================================="
			}
			Else
			{
				Line 1 "Name  Created                Protected # Users # Computers # Groups"
				Line 1 "==================================================================="
			}
		}
		If($HTML)
		{
			#V3.00 - pre-allocate rowdata
			## $rowdata = @()
			$rowdata  = New-Object System.Array[] $NumOUs
			$rowIndex = 0
		}

		ForEach($OU in $OUs)
		{
			$OUCount++
			$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing OU $($OU.CanonicalName) - OU # $OUCount of $NumOUs"
			
			#get counts of users, computers and groups in the OU
			
			[int]$UserCount = 0
			[int]$ComputerCount = 0
			[int]$GroupCount = 0
			$Results = @(Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
			$UserCount = $Results.Count

			$Results = @(Get-ADComputer -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
			$ComputerCount = $Results.Count

			$Results = @(Get-ADGroup -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
			$GroupCount = $Results.Count
			
			If($OU.ProtectedFromAccidentalDeletion -eq $False)
			{
				$UnprotectedOUs++
			}
			
			[string]$UserCountStr = "{0,7:N0}" -f $UserCount
			[string]$ComputerCountStr = "{0,11:N0}" -f $ComputerCount
			[string]$GroupCountStr = "{0,7:N0}" -f $GroupCount

			If($MSWord -or $PDF)
			{
				If($OU.ProtectedFromAccidentalDeletion -eq $True)
				{
					$Tmp = "Yes"
				}
				Else
				{
					$Tmp = "No"
				}

				$WordTableRowHash = @{ 
				OUName = $OUDisplayName; 
				OUCreated = $OU.Created.ToString(); 
				OUProtected = $Tmp;
				OUNumUsers = $UserCountStr;
				OUNumComputers = $ComputerCountStr;
				OUNumGroups = $GroupCountStr
				}

				## Add the hash to the array
				$ItemsWordTable.Add($WordTableRowHash) > $Null

				## Store "to highlight" cell references
				If($Tmp -eq "No") 
				{
					$WordHighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 3; }) > $Null
				}
				$CurrentServiceIndex++;
			}
			If($Text)
			{
				If($OU.ProtectedFromAccidentalDeletion -eq $True)
				{
					$tmp = "Yes"
				}
				Else
				{
					$tmp = "NO"
				}

				If(($OUDisplayName).Length -lt ($MaxOUNameLength))
				{
					[int]$NumOfSpaces = ($MaxOUNameLength * -1) 
				}
				Else
				{
					[int]$NumOfSpaces = -4
				}
				Line 1 ( "{0,$NumOfSpaces}  {1,-22} {2,-9} {3,-7} {4,-11}  {5,-7}" -f $OUDisplayName,$OU.Created.ToString(),$tmp,$UserCountStr,$ComputerCountStr,$GroupCountStr)
			}
			If($HTML)
			{
				If( $OU.ProtectedFromAccidentalDeletion -eq $True )
				{
					$Protected = 'Yes'
					$HTMLHighlightedCells = $htmlwhite
				}
				Else
				{
					$Protected = 'No'
					$HTMLHighlightedCells = $htmlred
				} 

				$rowData[ $rowIndex ] = @(
					$OUDisplayName,         $htmlwhite,
					$OU.Created.ToString(), $htmlwhite,
					$Protected,             $HTMLHighlightedCells,
					$UserCountStr,          ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
					$ComputerCountStr,      ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
					$GroupCountStr,         ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
				)
				$rowIndex++
			}
			$Results = $Null
			$UserCountStr = $Null
			$ComputerCountStr = $Null
			$GroupCountStr = $Null
		}
		
		If($MSWord -or $PDF)
		{
			## Add the table to the document, using the hashtable
			If($ItemsWordTable.Count -gt 0)	#3.07
			{
				$Table = AddWordTable -Hashtable $ItemsWordTable `
				-Columns OUName, OUCreated, OUProtected, OUNumUsers, OUNumComputers, OUNumGroups `
				-Headers "Name", "Created", "Protected", "# Users", "# Computers", "# Groups" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
				If($WordHighlightedCells.Count -gt 0)
				{
					SetWordCellFormat -Coordinates $WordHighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;
				}

				$Table.Columns.Item(1).Width = 125
				$Table.Columns.Item(2).Width = 100
				$Table.Columns.Item(3).Width = 55
				$Table.Columns.Item(4).Width = 55
				$Table.Columns.Item(5).Width = 70
				$Table.Columns.Item(6).Width = 55

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
				$Table = $Null
				$Results = $Null
				$UserCountStr = $Null
				$ComputerCountStr = $Null
				$GroupCountStr = $Null

				#added in V2.22
				If($UnprotectedOUs -gt 0)
				{
					WriteWordLine 0 0 "There are $($UnprotectedOUs) unprotected OUs out of $($NumOUs) OUs"
				}
			}
		}
		If($Text)
		{
			Line 0 ""
			If($UnprotectedOUs -gt 0)
			{
				Line 0 "There are $($UnprotectedOUs) unprotected OUs out of $($NumOUs) OUs"
				Line 0 ""
			}
		}
		If($HTML)
		{
			$columnWidths  = @( '225px', '300px', '55px', '55px', '75px', '55px' )
			$columnHeaders = @(
				'Name',        $htmlsb,
				'Created',     $htmlsb,
				'Protected',   $htmlsb,
				'# Users',     $htmlsb,
				'# Computers', $htmlsb,
				'# Groups',    $htmlsb
			)

			FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '765'
			#added in V2.22
			If($UnprotectedOUs -gt 0)
			{
				WriteHTMLLine 0 0 "There are $($UnprotectedOUs) unprotected OUs out of $($NumOUs) OUs"
			}

			$rowData = $null
		}
		
		$First = $False
	}
} ## end Function ProcessOrganizationalUnits
#endregion

#region Group information
Function ProcessGroupInformation
{
	## FIXME - v3.00 see optimizations applied to getDSUsers
	
	Write-Verbose "$(Get-Date -Format G): Writing group data"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Groups"
	}
	If($Text)
	{
		Line 0 "///  Groups  \\\"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Groups&nbsp;&nbsp;\\\"
	}

	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing groups in domain $($Domain)"
		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}
		
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt = "Domain $($Domain) (Forest Root)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
			}
			If($Text)
			{
				Line 0 "///  $($txt)  \\\"
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
			}
		}
		Else
		{
			$txt = "Domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
			}
			If($Text)
			{
				Line 0 "///  $($txt)  \\\"
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
			}
		}

		#get all Groups for the domain
		$Groups = $Null
		
		$Groups = Get-ADGroup -Filter * -Server $Domain -Properties Name, GroupCategory, GroupType -EA 0 | Sort-Object Name

		If( !$? )
		{
			$txt = "Could not retrieve Group data for domain $Domain, $( $error[ 0 ] )"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt '' $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue ## go to next domain
		}

		If( $null -eq $Groups )
		{
			$txt = "No Group data was retrieved for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt '' $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If($? -and $Null -ne $Groups)
		{
			#get counts
			
			Write-Verbose "$(Get-Date -Format G): `t`tGetting counts"
			
			[int]$SecurityCount = 0
			[int]$DistributionCount = 0
			[int]$GlobalCount = 0
			[int]$UniversalCount = 0
			[int]$DomainLocalCount = 0
			[int]$ContactsCount = 0
			[int]$GroupsWithSIDHistory = 0
			
			Write-Verbose "$(Get-Date -Format G): `t`t`tSecurity Groups"
			$Results = @($groups | Where-Object {$_.groupcategory -eq "Security"})
			
			[int]$SecurityCount = $Results.Count
			
			Write-Verbose "$(Get-Date -Format G): `t`t`tDistribution Groups"
			$Results = @($groups | Where-Object {$_.groupcategory -eq "Distribution"})
			
			[int]$DistributionCount = $Results.Count

			Write-Verbose "$(Get-Date -Format G): `t`t`tGlobal Groups"
			$Results = @($groups | Where-Object {$_.groupscope -eq "Global"})

			[int]$GlobalCount = $Results.Count

			Write-Verbose "$(Get-Date -Format G): `t`t`tUniversal Groups"
			$Results = @($groups | Where-Object {$_.groupscope -eq "Universal"})

			[int]$UniversalCount = $Results.Count
			
			Write-Verbose "$(Get-Date -Format G): `t`t`tDomain Local Groups"
			$Results = @($groups | Where-Object {$_.groupscope -eq "DomainLocal"})

			[int]$DomainLocalCount = $Results.Count

			Write-Verbose "$(Get-Date -Format G): `t`t`tGroups with SID History"
			$Results = $Null
			$Results = @( Get-ADObject -LDAPFilter "(sIDHistory=*)" -Server $Domain -Property objectClass, sIDHistory -EA 0 )
			$groups  = @( $Results | Where-Object {$_.objectClass -eq 'group'} )

			[int]$GroupsWithSIDHistory = $groups.Count

			Write-Verbose "$(Get-Date -Format G): `t`t`tContacts"
			$Results = $Null
			$Results = @(Get-ADObject -LDAPFilter "objectClass=Contact" -Server $Domain -EA 0)

			[int]$ContactsCount = $Results.Count

			[string]$TotalCountStr = "{0,7:N0}" -f ($SecurityCount + $DistributionCount)
			[string]$SecurityCountStr = "{0,7:N0}" -f $SecurityCount
			[string]$DomainLocalCountStr = "{0,7:N0}" -f $DomainLocalCount
			[string]$GlobalCountStr = "{0,7:N0}" -f $GlobalCount
			[string]$UniversalCountStr = "{0,7:N0}" -f $UniversalCount
			[string]$DistributionCountStr = "{0,7:N0}" -f $DistributionCount
			[string]$GroupsWithSIDHistoryStr = "{0,7:N0}" -f $GroupsWithSIDHistory
			[string]$ContactsCountStr = "{0,7:N0}" -f $ContactsCount
			
			Write-Verbose "$(Get-Date -Format G): `t`tBuild groups table"
			If($MSWORD -or $PDF)
			{
				$TableRange = $Script:doc.Application.Selection.Range
				[int]$Columns = 2
				[int]$Rows = 8
				$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
				$Table.Style = $Script:MyHash.Word_TableGrid
			
				$Table.Borders.InsideLineStyle = $wdLineStyleSingle
				$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
				$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(1,1).Range.Font.Bold = $True
				$Table.Cell(1,1).Range.Text = "Total Groups"
				$Table.Cell(1,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(1,2).Range.Text = $TotalCountStr
				$Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(2,1).Range.Font.Bold = $True
				$Table.Cell(2,1).Range.Text = "`tSecurity Groups"
				$Table.Cell(2,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(2,2).Range.Text = $SecurityCountStr
				$Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(3,1).Range.Font.Bold = $True
				$Table.Cell(3,1).Range.Text = "`t`tDomain Local"
				$Table.Cell(3,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(3,2).Range.Text = $DomainLocalCountStr
				$Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(4,1).Range.Font.Bold = $True
				$Table.Cell(4,1).Range.Text = "`t`tGlobal"
				$Table.Cell(4,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(4,2).Range.Text = $GlobalCountStr
				$Table.Cell(5,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(5,1).Range.Font.Bold = $True
				$Table.Cell(5,1).Range.Text = "`t`tUniversal"
				$Table.Cell(5,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(5,2).Range.Text = $UniversalCountStr
				$Table.Cell(6,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(6,1).Range.Font.Bold = $True
				$Table.Cell(6,1).Range.Text = "`tDistribution Groups"
				$Table.Cell(6,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(6,2).Range.Text = $DistributionCountStr
				$Table.Cell(7,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(7,1).Range.Font.Bold = $True
				$Table.Cell(7,1).Range.Text = "Groups with SID History"
				$Table.Cell(7,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(7,2).Range.Text = $GroupsWithSIDHistoryStr
				$Table.Cell(8,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(8,1).Range.Font.Bold = $True
				$Table.Cell(8,1).Range.Text = "Contacts"
				$Table.Cell(8,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(8,2).Range.Text = $ContactsCountStr

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
				$Table.AutoFitBehavior($wdAutoFitContent)

				#Return focus back to document
				$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
				$TableRange = $Null
				$Table = $Null
			}
			If($Text)
			{
				Line 1 "Total Groups`t`t`t: " $TotalCountStr
				Line 1 "`tSecurity Groups`t`t: " $SecurityCountStr
				Line 1 "`t`tDomain Local`t: " $DomainLocalCountStr
				Line 1 "`t`tGlobal`t`t: " $GlobalCountStr
				Line 1 "`t`tUniversal`t: " $UniversalCountStr
				Line 1 "`tDistribution Groups`t: " $DistributionCountStr
				Line 1 "Groups with SID History`t`t: " $GroupsWithSIDHistoryStr
				Line 1 "Contacts`t`t`t: " $ContactsCountStr
				Line 0 ""
			}
			If($HTML)
			{
				$columnHeaders = @("Total Groups",$htmlsb,$TotalCountStr,( $htmlwhite -bor $global:htmlAlignRight )) #3.06 add alignright

				$rowdata = New-Object System.Array[] 7
				$i = 0
				$rowdata[ $i++ ] = @("     Security Groups",$htmlsb,$SecurityCountStr,( $htmlwhite -bor $global:htmlAlignRight )) #3.06 add alignright
				$rowdata[ $i++ ] = @("          Domain Local",$htmlsb,$DomainLocalCountStr,( $htmlwhite -bor $global:htmlAlignRight )) #3.06 add alignright
				$rowdata[ $i++ ] = @("          Global",$htmlsb,$GlobalCountStr,( $htmlwhite -bor $global:htmlAlignRight )) #3.06 add alignright
				$rowdata[ $i++ ] = @("          Universal",$htmlsb,$UniversalCountStr,( $htmlwhite -bor $global:htmlAlignRight )) #3.06 add alignright
				$rowdata[ $i++ ] = @("     Distribution Groups",$htmlsb,$DistributionCountStr,( $htmlwhite -bor $global:htmlAlignRight )) #3.06 add alignright
				$rowdata[ $i++ ] = @("Groups with SID History",$htmlsb,$GroupsWithSIDHistoryStr,( $htmlwhite -bor $global:htmlAlignRight )) #3.06 add alignright
				$rowdata[ $i++ ] = @("Contacts",$htmlsb,$ContactsCountStr,( $htmlwhite -bor $global:htmlAlignRight )) #3.06 add alignright

				$columnWidths = @("250","75")
				FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "325"
				WriteHTMLLine 0 0 ''

				$rowData = $null
			}
			
			#get members of privileged groups
			$DomainInfo = $Null
			$DomainInfo = Get-ADDomain -Identity $Domain -EA 0
			
			If($? -and $Null -ne $DomainInfo)
			{
				$DomainAdminsSID = "$($DomainInfo.DomainSID)-512"
				$EnterpriseAdminsSID = "$($DomainInfo.DomainSID)-519"
				$SchemaAdminsSID = "$($DomainInfo.DomainSID)-518"
			}
			Else
			{
				$DomainAdminsSID = $Null
				$EnterpriseAdminsSID = $Null
				$SchemaAdminsSID = $Null
			}
			
			Write-Verbose "$(Get-Date -Format G): `t`tListing domain admins"
			$Admins = $Null
			
			try
			{
				$Admins = @(Get-ADGroupMember -Identity $DomainAdminsSID -Server $Domain -EA 0)
			}
			
			catch
			{
				#hide error
			}
			
			If($? -and $Null -ne $Admins)
			{
				[int]$AdminsCount = $Admins.Count
				$Admins = $Admins | Sort-Object Name
				[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
				
				If($MSWORD -or $PDF)
				{
					WriteWordLine 3 0 "Privileged Groups"
					WriteWordLine 4 0 "Domain Admins ($($AdminsCountStr) members):"
					$TableRange = $Script:doc.Application.Selection.Range
					[int]$Columns = 6
					[int]$Rows = $AdminsCount + 1
					[int]$xRow = 1
					$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.AutoFitBehavior($wdAutoFitFixed)
					$Table.Style = $Script:MyHash.Word_TableGrid
			
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "SamAccountName"
					$Table.Cell($xRow,3).Range.Font.Bold = $True
					$Table.Cell($xRow,3).Range.Text = "Domain"
					$Table.Cell($xRow,4).Range.Font.Bold = $True
					$Table.Cell($xRow,4).Range.Text = "Password Last Changed"
					$Table.Cell($xRow,5).Range.Font.Bold = $True
					$Table.Cell($xRow,5).Range.Text = "Password Never Expires"
					$Table.Cell($xRow,6).Range.Font.Bold = $True
					$Table.Cell($xRow,6).Range.Text = "Account Enabled"
				}
				If($Text)
				{
					Line 0 "Privileged Groups"
					Line 1 "Domain Admins ($AdminsCountStr members):"
					Line 2 "                                                                                                     Password    Password          "
					Line 2 "                                                                                                     Last        Never       Account"
					Line 2 "Name                                                SamAccountName        Domain                     Changed     Expires     Enabled"
					Line 2 "===================================================================================================================================="
					#       12345678901234567890123456789012345678901234567890SS12345678901234567890SS1234567890123456789012345SS1234567890SS1234567890SS12345
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Privileged Groups"
					WriteHTMLLine 4 0 "Domain Admins ($($AdminsCountStr) members):"
					#V3.00 pre-allocate rowdata
					## $rowdata = @()
					$rowData = New-Object System.Array[] $AdminsCount
					$rowIndx = 0
				}

				ForEach($Admin in $Admins)
				{
					If($MSWord -or $PDF)
					{
						$xRow++
					}
					
					#V3.01 change to use same method as EA & SA
					$dn = $Admin.distinguishedName
					$xServer = $dn.SubString( $dn.IndexOf( ',DC=' ) + 1 ).Replace( 'DC=', '' ).Replace( ',', '.' )

					If($Admin.ObjectClass -eq 'user')
					{
						$User = Get-ADUser -Identity $Admin.SID.value -Server $xServer `
						-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0
					}
					ElseIf($Admin.ObjectClass -eq 'group')
					{
						$User = Get-ADGroup -Identity $Admin.SID.value -Server $xServer -EA 0
					}
					Else
					{
						$User = $Null
					}

					If($? -and $Null -ne $User)
					{
						If($Admin.ObjectClass -eq 'user')
						{
							If($MSWord -or $PDF)
							{
								$Table.Cell($xRow,1).Range.Text = $User.Name
								$Table.Cell($xRow,2).Range.Text = $User.SamAccountName
								$Table.Cell($xRow,3).Range.Text = $xServer
								If($Null -eq $User.PasswordLastSet)
								{
									$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
									$Table.Cell($xRow,4).Range.Font.Bold  = $True
									$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
									$Table.Cell($xRow,4).Range.Text = "No Date Set"
								}
								Else
								{
									$Table.Cell($xRow,4).Range.Text = (Get-Date $User.PasswordLastSet -f d)
								}
								If($User.PasswordNeverExpires -eq $True)
								{
									$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorRed
									$Table.Cell($xRow,5).Range.Font.Bold  = $True
									$Table.Cell($xRow,5).Range.Font.Color = $WDColorBlack
								}
								$Table.Cell($xRow,5).Range.Text = $User.PasswordNeverExpires.ToString()
								If($User.Enabled -eq $False)
								{
									$Table.Cell($xRow,6).Shading.BackgroundPatternColor = $wdColorRed
									$Table.Cell($xRow,6).Range.Font.Bold  = $True
									$Table.Cell($xRow,6).Range.Font.Color = $WDColorBlack
								}
								$Table.Cell($xRow,6).Range.Text = $User.Enabled.ToString()
							}
							If($Text)
							{
								If($Null -eq $User.PasswordLastSet)
								{
									$PasswordLastSet = "No Date Set"
								}
								Else
								{
									$PasswordLastSet = (Get-Date $User.PasswordLastSet -f d)
								}
								Line 2 ( "{0,-50}  {1,-20}  {2,-25}  {3,-10}  {4,-10}  {5,-5}" -f $User.Name,$User.SamAccountName,$xServer,$PasswordLastSet,$User.PasswordNeverExpires.ToString(),$User.Enabled.ToString())
							}
							If($HTML)
							{
								$UserName = $User.Name
								$UserSamAccountName = $User.SamAccountName
								$Domain = $xServer
								If($Null -eq $User.PasswordLastSet)
								{
									$PasswordLastSet = "No Date Set"
								}
								Else
								{
									$PasswordLastSet = (Get-Date $User.PasswordLastSet -f d)
								}
								$PasswordNeverExpires = $User.PasswordNeverExpires.ToString()
								$Enabled = $User.Enabled.ToString()
							}
						}
						ElseIf($Admin.ObjectClass -eq 'group')
						{
							If($MSWord -or $PDF)
							{
								$Table.Cell($xRow,1).Range.Text = "$($User.Name) (group)"
								$Table.Cell($xRow,2).Range.Text = "$($User.SamAccountName)"
								$Table.Cell($xRow,3).Range.Text = $xServer
								$Table.Cell($xRow,4).Range.Text = "N/A"
								$Table.Cell($xRow,5).Range.Text = "N/A"
								$Table.Cell($xRow,6).Range.Text = "N/A"
							}
							If($Text)
							{
								Line 2 ( "{0,-43} (group) {1,-20}  {2,-25}  {3,-10}  {4,-10}  {5,-5}" -f $User.Name,$User.SamAccountName,$xServer,"N/A","N/A","N/A")
							}
							If($HTML)
							{
								$UserName = "$($User.Name) (group)"
								$UserSamAccountName = "$($User.SamAccountName)"
								$Domain = $xServer
								$PasswordLastSet = "N/A"
								$PasswordNeverExpires = "N/A"
								$Enabled = "N/A"
							}
						}
					}
					Else
					{
						If($MSWord -or $PDF)
						{
							$Table.Cell($xRow,1).Range.Text = $Admin.SID.Value
							$Table.Cell($xRow,2).Range.Text = "Unknown"
							$Table.Cell($xRow,3).Range.Text = $xServer
							$Table.Cell($xRow,4).Range.Text = "Unknown"
							$Table.Cell($xRow,5).Range.Text = "Unknown"
							$Table.Cell($xRow,6).Range.Text = "Unknown"
						}
						If($Text)
						{
							Line 2 ( "{0,-50}  {1,-20}  {2,-25}  {3,-10}  {4,-10}  {5,-5}" -f $Admin.SID.Value,"Unknown",$xServer,"Unknown","Unknown","Unknown")
						}
						If($HTML)
						{
							$UserName = $Admin.SID.Value
							$UserSamAccountName = "Unknown"
							$Domain = $xServer
							$PasswordLastSet = "Unknown"
							$PasswordNeverExpires = "Unknown"
							$Enabled = "Unknown"
						}
					}
					
					If($HTML)
					{
						If($PasswordLastSet -eq "No Date Set")
						{
							$HTMLHighlightedCells1 = $htmlred
						}
						Else
						{
							$HTMLHighlightedCells1 = $htmlwhite
						}
						
						If($PasswordNeverExpires -eq "True")
						{
							$HTMLHighlightedCells2 = $htmlred
						}
						Else
						{
							$HTMLHighlightedCells2 = $htmlwhite
						}
						
						If($Enabled -eq "False")
						{
							$HTMLHighlightedCells3 = $htmlred
						}
						Else
						{
							$HTMLHighlightedCells3 = $htmlwhite
						}
						
						$rowdata[ $rowIndx ] = @(
							$UserName,             $htmlwhite,
							$UserSamAccountName,   $htmlwhite,
							$Domain,               $htmlwhite,
							$PasswordLastSet,      $HTMLHighlightedCells1,
							$PasswordNeverExpires, $HTMLHighlightedCells2,
							$Enabled,              $HTMLHighlightedCells3
						)
						$rowIndx++
					}
				}
			
				If($MSWord -or $PDF)
				{
					#set column widths
					$xcols = $table.columns

					ForEach($xcol in $xcols)
					{
						Switch ($xcol.Index)
						{
							1 {$xcol.width = 100; Break}
							2 {$xcol.width = 100; Break}
							3 {$xcol.width = 108; Break}
							4 {$xcol.width = 66; Break}
							5 {$xcol.width = 56; Break}
							6 {$xcol.width = 56; Break}
						}
					}

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
					$Table.AutoFitBehavior($wdAutoFitFixed)

					#Return focus back to document
					$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
					$TableRange = $Null
					$Table = $Null
				}
				If($Text)
				{
					Line 0 ""
				}
				If($HTML)
				{
					$columnWidths  = @( '125px', '125px', '135px', '70px', '60px', '60px' )
					$columnHeaders = @(
						'Name',                   $htmlsb,
						'SamAccountName',         $htmlsb,
						'Domain',                 $htmlsb,
						'Password Last Changed',  $htmlsb,
						'Password Never Expires', $htmlsb,
						'Account Enabled',        $htmlsb
					)

					FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '575'
					WriteHTMLLine 0 0 ''

					$rowData = $null
				}
			}
			ElseIf(!$?)
			{
				$txt = "Unable to retrieve Domain Admins group membership"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt "" $Null 0 $False $True
				}
				If($Text)
				{
					Line 0 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}
			}
			Else
			{
				$txt = "Domain Admins: <None>"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 4 0 $txt
				}
				If($Text)
				{
					Line 0 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Domain Admins: None"
				}
			}

			If($Domain -eq $Script:ForestRootDomain)
			{
				Write-Verbose "$(Get-Date -Format G): `t`tListing enterprise admins"
				$Admins = $Null
			
				try
				{
					$Admins = @(Get-ADGroupMember -Identity $EnterpriseAdminsSID -Server $Domain -EA 0)
				}
				
				catch
				{
					#hide error
				}
				
				If($? -and $Null -ne $Admins)
				{
					[int]$AdminsCount = $Admins.Count
					$Admins = $Admins | Sort-Object Name
					[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
					
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 "Enterprise Admins ($($AdminsCountStr) members):"
						$TableRange = $Script:doc.Application.Selection.Range
						[int]$Columns = 6
						[int]$Rows = $AdminsCount + 1
						[int]$xRow = 1
						$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
						$Table.AutoFitBehavior($wdAutoFitFixed)
						$Table.Style = $Script:MyHash.Word_TableGrid
				
						$Table.rows.first.headingformat = $wdHeadingFormatTrue
						$Table.Borders.InsideLineStyle = $wdLineStyleSingle
						$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

						$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,1).Range.Font.Bold = $True
						$Table.Cell($xRow,1).Range.Text = "Name"
						$Table.Cell($xRow,2).Range.Font.Bold = $True
						$Table.Cell($xRow,2).Range.Text = "SamAccountName"
						$Table.Cell($xRow,3).Range.Font.Bold = $True
						$Table.Cell($xRow,3).Range.Text = "Domain"
						$Table.Cell($xRow,4).Range.Font.Bold = $True
						$Table.Cell($xRow,4).Range.Text = "Password Last Changed"
						$Table.Cell($xRow,5).Range.Font.Bold = $True
						$Table.Cell($xRow,5).Range.Text = "Password Never Expires"
						$Table.Cell($xRow,6).Range.Font.Bold = $True
						$Table.Cell($xRow,6).Range.Text = "Account Enabled"
					}
					If($Text)
					{
						Line 1 "Enterprise Admins ($AdminsCountStr members):"
						Line 2 "                                                                                                     Password    Password          "
						Line 2 "                                                                                                     Last        Never       Account"
						Line 2 "Name                                                SamAccountName        Domain                     Changed     Expires     Enabled"
						Line 2 "===================================================================================================================================="
						#       12345678901234567890123456789012345678901234567890SS12345678901234567890SS1234567890123456789012345SS1234567890SS1234567890SS12345
					}
					If($HTML)
					{
						WriteHTMLLine 4 0 "Enterprise Admins ($($AdminsCountStr) members):"
						#V3.00 pre-allocate rowdata
						## $rowdata = @()
						$rowData = New-Object System.Array[] $AdminsCount
						$rowIndx = 0
					}

					ForEach($Admin in $Admins)
					{
						If($MSWord -or $PDF)
						{
							$xRow++
						}
						#V3.00 - 3-to-1 speed advantage, new code
						## FIXME - apply this speed-up to other code paths - MBS
						$dn = $Admin.distinguishedName
						$xServer = $dn.SubString( $dn.IndexOf( ',DC=' ) + 1 ).Replace( 'DC=', '' ).Replace( ',', '.' )

						If($Admin.ObjectClass -eq 'user')
						{
							$User = Get-ADUser -Identity $Admin.SID.value -Server $xServer `
							-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0
						}
						ElseIf($Admin.ObjectClass -eq 'group')
						{
							$User = Get-ADGroup -Identity $Admin.SID.value -Server $xServer -EA 0
						}
						Else
						{
							$User = $Null
						}
						
						If($? -and $Null -ne $User)
						{
							If($Admin.ObjectClass -eq 'user')
							{
								If($MSWord -or $PDF)
								{
									$Table.Cell($xRow,1).Range.Text = $User.Name
									$Table.Cell($xRow,2).Range.Text = $User.SamAccountName
									$Table.Cell($xRow,3).Range.Text = $xServer
									If($Null -eq $User.PasswordLastSet)
									{
										$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,4).Range.Font.Bold  = $True
										$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
										$Table.Cell($xRow,4).Range.Text = "No Date Set"
									}
									Else
									{
										$Table.Cell($xRow,4).Range.Text = (Get-Date $User.PasswordLastSet -f d)
									}
									If($User.PasswordNeverExpires -eq $True)
									{
										$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,5).Range.Font.Bold  = $True
										$Table.Cell($xRow,5).Range.Font.Color = $WDColorBlack
									}
									$Table.Cell($xRow,5).Range.Text = $User.PasswordNeverExpires.ToString()
									If($User.Enabled -eq $False)
									{
										$Table.Cell($xRow,6).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,6).Range.Font.Bold  = $True
										$Table.Cell($xRow,6).Range.Font.Color = $WDColorBlack
									}
									$Table.Cell($xRow,6).Range.Text = $User.Enabled.ToString()
								}
								If($Text)
								{
									If($Null -eq $User.PasswordLastSet)
									{
										$PasswordLastSet = "No Date Set"
									}
									Else
									{
										$PasswordLastSet = (Get-Date $User.PasswordLastSet -f d)
									}
									Line 2 ( "{0,-50}  {1,-20}  {2,-25}  {3,-10}  {4,-10}  {5,-5}" -f $User.Name,$User.SamAccountName,$xServer,$PasswordLastSet,$User.PasswordNeverExpires.ToString(),$User.Enabled.ToString())
								}
								If($HTML)
								{
									$UserName = $User.Name
									$UserSamAccountName = $User.SamAccountName
									$Domain = $xServer
									If($Null -eq $User.PasswordLastSet)
									{
										$PasswordLastSet = "No Date Set"
									}
									Else
									{
										$PasswordLastSet = (Get-Date $User.PasswordLastSet -f d)
									}
									$PasswordNeverExpires = $User.PasswordNeverExpires.ToString()
									$Enabled = $User.Enabled.ToString()
								}
							}
							ElseIf($Admin.ObjectClass -eq 'group')
							{
								If($MSWord -or $PDF)
								{
									$Table.Cell($xRow,1).Range.Text = "$($User.Name) (group)"
									$Table.Cell($xRow,2).Range.Text = "$($User.SamAccountName)"
									$Table.Cell($xRow,3).Range.Text = $xServer
									$Table.Cell($xRow,4).Range.Text = "N/A"
									$Table.Cell($xRow,5).Range.Text = "N/A"
									$Table.Cell($xRow,6).Range.Text = "N/A"
								}
								If($Text)
								{
									Line 2 ( "{0,-43} (group) {1,-20}  {2,-25}  {3,-10}  {4,-10}  {5,-5}" -f $User.Name,$User.SamAccountName,$xServer,"N/A","N/A","N/A")
								}
								If($HTML)
								{
									$UserName = "$($User.Name) (group)"
									$UserSamAccountName = "$($User.SamAccountName)"
									$Domain = $xServer
									$PasswordLastSet = "N/A"
									$PasswordNeverExpires = "N/A"
									$Enabled = "N/A"
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$Table.Cell($xRow,1).Range.Text = $Admin.SID.Value
								$Table.Cell($xRow,2).Range.Text = "Unknown"
								$Table.Cell($xRow,3).Range.Text = $xServer
								$Table.Cell($xRow,4).Range.Text = "Unknown"
								$Table.Cell($xRow,5).Range.Text = "Unknown"
								$Table.Cell($xRow,6).Range.Text = "Unknown"
							}
							If($Text)
							{
								Line 2 ( "{0,-50}  {1,-20}  {2,-25}  {3,-10}  {4,-10}  {5,-5}" -f $Admin.SID.Value,"Unknown",$xServer,"Unknown","Unknown","Unknown")
							}
							If($HTML)
							{
								$UserName = $Admin.SID.Value
								$UserSamAccountName = "Unknown"
								$Domain = $xServer
								$PasswordLastSet = "Unknown"
								$PasswordNeverExpires = "Unknown"
								$Enabled = "Unknown"
							}
						}
						
						If($HTML)
						{
							If($PasswordLastSet -eq "No Date Set")
							{
								$HTMLHighlightedCells1 = $htmlred
							}
							Else
							{
								$HTMLHighlightedCells1 = $htmlwhite
							}
							
							If($PasswordNeverExpires -eq "True")
							{
								$HTMLHighlightedCells2 = $htmlred
							}
							Else
							{
								$HTMLHighlightedCells2 = $htmlwhite
							}
							
							If($Enabled -eq "False")
							{
								$HTMLHighlightedCells3 = $htmlred
							}
							Else
							{
								$HTMLHighlightedCells3 = $htmlwhite
							}
							
							$rowdata[ $rowIndx ] = @(
								$UserName,             $htmlwhite,
								$UserSamAccountName,   $htmlwhite,
								$Domain,               $htmlwhite,
								$PasswordLastSet,      $HTMLHighlightedCells1,
								$PasswordNeverExpires, $HTMLHighlightedCells2,
								$Enabled,              $HTMLHighlightedCells3
							)
							$rowIndx++
						}
					}
				
					If($MSWord -or $PDF)
					{
						#set column widths
						$xcols = $table.columns

						ForEach($xcol in $xcols)
						{
							Switch ($xcol.Index)
							{
								1 {$xcol.width = 100; Break}
								2 {$xcol.width = 100; Break}
								3 {$xcol.width = 108; Break}
								4 {$xcol.width = 66; Break}
								5 {$xcol.width = 56; Break}
								6 {$xcol.width = 56; Break}
							}
						}

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
						$Table.AutoFitBehavior($wdAutoFitFixed)

						#Return focus back to document
						$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

						#move to the end of the current document
						$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
						$TableRange = $Null
						$Table = $Null
					}
					If($Text)
					{
						Line 0 ""
					}
					If($HTML)
					{
						$columnWidths  = @( '125px', '125px', '135px', '70px', '60px', '60px' )
						$columnHeaders = @(
							'Name',                   $htmlsb,
							'SamAccountName',         $htmlsb,
							'Domain',                 $htmlsb,
							'Password Last Changed',  $htmlsb,
							'Password Never Expires', $htmlsb,
							'Account Enabled',        $htmlsb
						)

						FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '575'
						WriteHTMLLine 0 0 ''

						$rowData = $null
					}
				}
				ElseIf(!$?)
				{
					$txt1 = "Enterprise Admins:"
					$txt2 = "Unable to retrieve Enterprise Admins group membership"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 $txt1
						WriteWordLine 0 0 $txt2 "" $Null 0 $False $True
					}
					If($Text)
					{
						Line 0 $txt1
						Line 0 $txt2
					}
					If($HTML)
					{
						WriteHTMLLine 4 0 $txt1
						WriteHTMLLine 0 0 $txt2
					}
				}
				Else
				{
					$txt1 = "Enterprise Admins:"
					$txt2 = "<None>"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 $txt1
						WriteWordLine 0 0 $txt2
					}
					If($Text)
					{
						Line 0 $txt1
						Line 0 $txt2
					}
					If($HTML)
					{
						WriteHTMLLine 4 0 $txt1
						WriteHTMLLine 0 0 "None"
					}
				}
			}
			
			If($Domain -eq $Script:ForestRootDomain)
			{
				Write-Verbose "$(Get-Date -Format G): `t`tListing schema admins"
				$Admins = $Null
			
				try
				{
					$Admins = @(Get-ADGroupMember -Identity $SchemaAdminsSID -Server $Domain -EA 0)
				}
				
				catch
				{
					#hide error
				}
				
				If($? -and $Null -ne $Admins)
				{
					[int]$AdminsCount = $Admins.Count
					$Admins = $Admins | Sort-Object Name
					[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
					
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 "Schema Admins ($($AdminsCountStr) members): "
						$TableRange = $Script:doc.Application.Selection.Range
						[int]$Columns = 6
						[int]$Rows = $AdminsCount + 1
						[int]$xRow = 1
						$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
						$Table.AutoFitBehavior($wdAutoFitFixed)
						$Table.Style = $Script:MyHash.Word_TableGrid
				
						$Table.rows.first.headingformat = $wdHeadingFormatTrue
						$Table.Borders.InsideLineStyle = $wdLineStyleSingle
						$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

						$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,1).Range.Font.Bold = $True
						$Table.Cell($xRow,1).Range.Text = "Name"
						$Table.Cell($xRow,2).Range.Font.Bold = $True
						$Table.Cell($xRow,2).Range.Text = "SamAccountName"
						$Table.Cell($xRow,3).Range.Font.Bold = $True
						$Table.Cell($xRow,3).Range.Text = "Domain"
						$Table.Cell($xRow,4).Range.Font.Bold = $True
						$Table.Cell($xRow,4).Range.Text = "Password Last Changed"
						$Table.Cell($xRow,5).Range.Font.Bold = $True
						$Table.Cell($xRow,5).Range.Text = "Password Never Expires"
						$Table.Cell($xRow,6).Range.Font.Bold = $True
						$Table.Cell($xRow,6).Range.Text = "Account Enabled"
					}
					If($Text)
					{
						Line 1 "Schema Admins ($($AdminsCountStr) members): "
						Line 2 "                                                                                                     Password    Password          "
						Line 2 "                                                                                                     Last        Never       Account"
						Line 2 "Name                                                SamAccountName        Domain                     Changed     Expires     Enabled"
						Line 2 "===================================================================================================================================="
						#       12345678901234567890123456789012345678901234567890SS12345678901234567890SS1234567890123456789012345SS1234567890SS1234567890SS12345
					}
					If($HTML)
					{
						#V3.00 pre-allocate rowdata
						## $rowdata = @()
						WriteHTMLLine 4 0 "Schema Admins ($($AdminsCountStr) members): "
						$rowData = New-Object System.Array[] $AdminsCount
						$rowIndx = 0
					}

					ForEach($Admin in $Admins)
					{
						If($MSWord -or $PDF)
						{
							$xRow++
						}

						$dn = $Admin.distinguishedName
						$xServer = $dn.SubString( $dn.IndexOf( ',DC=' ) + 1 ).Replace( 'DC=', '' ).Replace( ',', '.' )

						If($Admin.ObjectClass -eq 'user')
						{
							$User = Get-ADUser -Identity $Admin.SID.value -Server $xServer `
							-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0 
						}
						ElseIf($Admin.ObjectClass -eq 'group')
						{
							$User = Get-ADGroup -Identity $Admin.SID.value -Server $xServer -EA 0 
						}
						Else
						{
							$User = $Null
						}
						
						If($? -and $Null -ne $User)
						{
							If($Admin.ObjectClass -eq 'user')
							{
								If($MSWord -or $PDF)
								{
									$Table.Cell($xRow,1).Range.Text = $User.Name
									$Table.Cell($xRow,2).Range.Text = $User.SamAccountName
									$Table.Cell($xRow,3).Range.Text = $xServer
									If($Null -eq $User.PasswordLastSet)
									{
										$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,4).Range.Font.Bold  = $True
										$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
										$Table.Cell($xRow,4).Range.Text = "No Date Set"
									}
									Else
									{
										$Table.Cell($xRow,4).Range.Text = (Get-Date $User.PasswordLastSet -f d)
									}
									If($User.PasswordNeverExpires -eq $True)
									{
										$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,5).Range.Font.Bold  = $True
										$Table.Cell($xRow,5).Range.Font.Color = $WDColorBlack
									}
									$Table.Cell($xRow,5).Range.Text = $User.PasswordNeverExpires.ToString()
									If($User.Enabled -eq $False)
									{
										$Table.Cell($xRow,6).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,6).Range.Font.Bold  = $True
										$Table.Cell($xRow,6).Range.Font.Color = $WDColorBlack
									}
									$Table.Cell($xRow,6).Range.Text = $User.Enabled.ToString()
								}
								If($Text)
								{
									If($Null -eq $User.PasswordLastSet)
									{
										$PasswordLastSet = "No Date Set"
									}
									Else
									{
										$PasswordLastSet = (Get-Date $User.PasswordLastSet -f d)
									}
									Line 2 ( "{0,-50}  {1,-20}  {2,-25}  {3,-10}  {4,-10}  {5,-5}" -f $User.Name,$User.SamAccountName,$xServer,$PasswordLastSet,$User.PasswordNeverExpires.ToString(),$User.Enabled.ToString())
								}
								If($HTML)
								{
									$UserName = $User.Name
									$UserSamAccountName = $User.SamAccountName
									$Domain = $xServer
									If($Null -eq $User.PasswordLastSet)
									{
										$PasswordLastSet = "No Date Set"
									}
									Else
									{
										$PasswordLastSet = (Get-Date $User.PasswordLastSet -f d)
									}
									$PasswordNeverExpires = $User.PasswordNeverExpires.ToString()
									$Enabled = $User.Enabled.ToString()
								}
							}
							ElseIf($Admin.ObjectClass -eq 'group')
							{
								If($MSWord -or $PDF)
								{
									$Table.Cell($xRow,1).Range.Text = "$($User.Name) (group)"
									$Table.Cell($xRow,2).Range.Text = "$($User.SamAccountName)"
									$Table.Cell($xRow,3).Range.Text = $xServer
									$Table.Cell($xRow,4).Range.Text = "N/A"
									$Table.Cell($xRow,5).Range.Text = "N/A"
									$Table.Cell($xRow,6).Range.Text = "N/A"
								}
								If($Text)
								{
									Line 2 ( "{0,-43} (group) {1,-20}  {2,-25}  {3,-10}  {4,-10}  {5,-5}" -f $User.Name,$User.SamAccountName,$xServer,"N/A","N/A","N/A")
								}
								If($HTML)
								{
									$UserName = "$($User.Name) (group)"
									$UserSamAccountName = "$($User.SamAccountName)"
									$Domain = $xServer
									$PasswordLastSet = "N/A"
									$PasswordNeverExpires = "N/A"
									$Enabled = "N/A"
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$Table.Cell($xRow,1).Range.Text = $Admin.SID.Value
								$Table.Cell($xRow,2).Range.Text = "Unknown"
								$Table.Cell($xRow,3).Range.Text = $xServer
								$Table.Cell($xRow,4).Range.Text = "Unknown"
								$Table.Cell($xRow,5).Range.Text = "Unknown"
								$Table.Cell($xRow,6).Range.Text = "Unknown"
							}
							If($Text)
							{
								Line 2 ( "{0,-50}  {1,-20}  {2,-25}  {3,-10}  {4,-10}  {5,-5}" -f $Admin.SID.Value,"Unknown",$xServer,"Unknown","Unknown","Unknown")
							}
							If($HTML)
							{
								$UserName = $Admin.SID.Value
								$UserSamAccountName = "Unknown"
								$Domain = $xServer
								$PasswordLastSet = "Unknown"
								$PasswordNeverExpires = "Unknown"
								$Enabled = "Unknown"
							}
						}
						If($HTML)
						{
							If($PasswordLastSet -eq "No Date Set")
							{
								$HTMLHighlightedCells1 = $htmlred
							}
							Else
							{
								$HTMLHighlightedCells1 = $htmlwhite
							}
							
							If($PasswordNeverExpires -eq "True")
							{
								$HTMLHighlightedCells2 = $htmlred
							}
							Else
							{
								$HTMLHighlightedCells2 = $htmlwhite
							}
							
							If($Enabled -eq "False")
							{
								$HTMLHighlightedCells3 = $htmlred
							}
							Else
							{
								$HTMLHighlightedCells3 = $htmlwhite
							}
							
							$rowdata[ $rowIndx ] = @(
								$UserName,             $htmlwhite,
								$UserSamAccountName,   $htmlwhite,
								$Domain,               $htmlwhite,
								$PasswordLastSet,      $HTMLHighlightedCells1,
								$PasswordNeverExpires, $HTMLHighlightedCells2,
								$Enabled,              $HTMLHighlightedCells3
							)
							$rowIndx++
						}
					}
				
					If($MSWord -or $PDF)
					{
						#set column widths
						$xcols = $table.columns

						ForEach($xcol in $xcols)
						{
							Switch ($xcol.Index)
							{
								1 {$xcol.width = 100; Break}
								2 {$xcol.width = 100; Break}
								3 {$xcol.width = 108; Break}
								4 {$xcol.width = 66; Break}
								5 {$xcol.width = 56; Break}
								6 {$xcol.width = 56; Break}
							}
						}
						
						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
						$Table.AutoFitBehavior($wdAutoFitFixed)

						#Return focus back to document
						$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

						#move to the end of the current document
						$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
						$TableRange = $Null
						$Table = $Null
					}
					If($Text)
					{
						Line 0 ""
					}
					If($HTML)
					{
						$columnWidths  = @( '125px', '125px', '135px', '70px', '60px', '60px' )
						$columnHeaders = @(
							'Name',                   $htmlsb,
							'SamAccountName',         $htmlsb,
							'Domain',                 $htmlsb,
							'Password Last Changed',  $htmlsb,
							'Password Never Expires', $htmlsb,
							'Account Enabled',        $htmlsb
						)

						FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '575'
						WriteHTMLLine 0 0 ''

						$rowdata = $null
					}
				}
				ElseIf(!$?)
				{
					$txt1 = "Schema Admins: "
					$txt2 = "Unable to retrieve Schema Admins group membership"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 $txt1
						WriteWordLine 0 0 $txt2 "" $Null 0 $False $True
					}
					If($Text)
					{
						Line 0 $txt1
						Line 0 $txt2
					}
					If($HTML)
					{
						WriteHTMLLine 4 0 $txt1
						WriteHTMLLine 0 0 $txt2
					}
				}
				Else
				{
					$txt1 = "Schema Admins: "
					$txt2 = "<None>"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 $txt1
						WriteWordLine 0 0 $txt2
					}
					If($Text)
					{
						Line 0 $txt1
						Line 0 $txt2
					}
					If($HTML)
					{
						WriteHTMLLine 4 0 $txt1
						WriteHTMLLine 0 0 "None"
					}
				}
			}

			Write-Verbose "$(Get-Date -Format G): `t`tListing users with AdminCount=1"
			$AdminCounts = @(Get-ADUser -LDAPFilter "(admincount=1)"  -Server $Domain -EA 0)
			
			If($? -and $Null -ne $AdminCounts)
			{
				$AdminCounts = $AdminCounts | Sort-Object Name
				[int]$AdminsCount = $AdminCounts.Count
				[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
				
				If($MSWORD -or $PDF)
				{
					WriteWordLine 4 0 "Users with AdminCount=1 ($AdminsCountStr users):"
					$TableRange = $Script:doc.Application.Selection.Range
					[int]$Columns = 6
					[int]$Rows = $AdminsCount + 1
					[int]$xRow = 1
					$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.AutoFitBehavior($wdAutoFitFixed)
					$Table.Style = $Script:MyHash.Word_TableGrid
			
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "SamAccountName"
					$Table.Cell($xRow,3).Range.Font.Bold = $True
					$Table.Cell($xRow,3).Range.Text = "Domain"
					$Table.Cell($xRow,4).Range.Font.Bold = $True
					$Table.Cell($xRow,4).Range.Text = "Password Last Changed"
					$Table.Cell($xRow,5).Range.Font.Bold = $True
					$Table.Cell($xRow,5).Range.Text = "Password Never Expires"
					$Table.Cell($xRow,6).Range.Font.Bold = $True
					$Table.Cell($xRow,6).Range.Text = "Account Enabled"
				}
				If($Text)
				{
					Line 1 "Users with AdminCount=1 ($AdminsCountStr users):"
					Line 2 "                                                                                                     Password     Password           "
					Line 2 "                                                                                                     Last         Never       Account"
					Line 2 "Name                                                SamAccountName        Domain                     Changed      Expires     Enabled"
					Line 2 "====================================================================================================================================="
					#       12345678901234567890123456789012345678901234567890SS12345678901234567890SS1234567890123456789012345SS12345678901SS1234567890SS12345
					#                                                                                                            No Date Set
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Users with AdminCount=1 ($($AdminsCountStr) users):"
					#V3.00 pre-allocate rowdata
					## $rowdata = @()
					$rowData = New-Object System.Array[] $AdminsCount
					$rowIndx = 0
				}

				ForEach($Admin in $AdminCounts)
				{
					$User = Get-ADUser -Identity $Admin.SID -Server $Domain `
					-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0 
					
					If($MSWord -or $PDF)
					{
						$xRow++
					}
					
					If($? -and $Null -ne $User)
					{
						If($MSWord -or $PDF)
						{
							$Table.Cell($xRow,1).Range.Text = $User.Name
							$Table.Cell($xRow,2).Range.Text = $User.SamAccountName
							$Table.Cell($xRow,3).Range.Text = $xServer
							If($Null -eq $User.PasswordLastSet)
							{
								$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
								$Table.Cell($xRow,4).Range.Font.Bold  = $True
								$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
								$Table.Cell($xRow,4).Range.Text = "No Date Set"
							}
							Else
							{
								$Table.Cell($xRow,4).Range.Text = (Get-Date $User.PasswordLastSet -f d)
							}
							If($User.PasswordNeverExpires -eq $True)
							{
								$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorRed
								$Table.Cell($xRow,5).Range.Font.Bold  = $True
								$Table.Cell($xRow,5).Range.Font.Color = $WDColorBlack
							}
							$Table.Cell($xRow,5).Range.Text = $User.PasswordNeverExpires.ToString()
							If($User.Enabled -eq $False)
							{
								$Table.Cell($xRow,6).Shading.BackgroundPatternColor = $wdColorRed
								$Table.Cell($xRow,6).Range.Font.Bold  = $True
								$Table.Cell($xRow,6).Range.Font.Color = $WDColorBlack
							}
							$Table.Cell($xRow,6).Range.Text = $User.Enabled.ToString()
						}
						If($Text)
						{
							If($Null -eq $User.PasswordLastSet)
							{
								$PasswordLastSet = "No Date Set"
							}
							Else
							{
								$PasswordLastSet = (Get-Date $User.PasswordLastSet -f d)
							}
							#V3.00
							#$UserEnabled = $User.Enabled.ToString()
							Line 2 ( "{0,-50}  {1,-20}  {2,-25}  {3,-11}  {4,-10}  {5,-5}" -f $User.Name,$User.SamAccountName,$xServer,$PasswordLastSet,$User.PasswordNeverExpires.ToString(),$User.Enabled.ToString())
						}
						If($HTML)
						{
							$UserName = $User.Name
							$UserSamAccountName = $User.SamAccountName
							$Domain = $xServer
							If($Null -eq $User.PasswordLastSet)
							{
								$PasswordLastSet = "No Date Set"
							}
							Else
							{
								$PasswordLastSet = (Get-Date $User.PasswordLastSet -f d)
							}
							$PasswordNeverExpires = $User.PasswordNeverExpires.ToString()
							$Enabled = $User.Enabled.ToString()
						}
					}
					Else
					{
						If($MSWord -or $PDF)
						{
							$Table.Cell($xRow,1).Range.Text = $Admin.SID.Value
							$Table.Cell($xRow,2).Range.Text = "Unknown"
							$Table.Cell($xRow,3).Range.Text = $xServer
							$Table.Cell($xRow,4).Range.Text = "Unknown"
							$Table.Cell($xRow,5).Range.Text = "Unknown"
							$Table.Cell($xRow,6).Range.Text = "Unknown"
						}
						If($Text)
						{
							Line 2 ( "{0,-50}  {1,-20}  {2,-25}  {3,-10}  {4,-10}  {5,-5}" -f $Admin.SID.Value,"Unknown",$xServer,"Unknown","Unknown","Unknown")
						}
						If($HTML)
						{
							$UserName = $Admin.SID.Value
							$UserSamAccountName = "Unknown"
							$Domain = $xServer
							$PasswordLastSet = "Unknown"
							$PasswordNeverExpires = "Unknown"
							$Enabled = "Unknown"
						}
					}
					If($HTML)
					{
						If($PasswordLastSet -eq "No Date Set")
						{
							$HTMLHighlightedCells1 = $htmlred
						}
						Else
						{
							$HTMLHighlightedCells1 = $htmlwhite
						}
						
						If($PasswordNeverExpires -eq "True")
						{
							$HTMLHighlightedCells2 = $htmlred
						}
						Else
						{
							$HTMLHighlightedCells2 = $htmlwhite
						}
						
						If($Enabled -eq "False")
						{
							$HTMLHighlightedCells3 = $htmlred
						}
						Else
						{
							$HTMLHighlightedCells3 = $htmlwhite
						}
						
						$rowdata[ $rowIndx ] = @(
							$UserName,             $htmlwhite,
							$UserSamAccountName,   $htmlwhite,
							$Domain,               $htmlwhite,
							$PasswordLastSet,      $HTMLHighlightedCells1,
							$PasswordNeverExpires, $HTMLHighlightedCells2,
							$Enabled,              $HTMLHighlightedCells3
						)
						$rowIndx++
					}
				}
				
				If($MSWord -or $PDF)
				{
					#set column widths
					$xcols = $table.columns

					ForEach($xcol in $xcols)
					{
						Switch ($xcol.Index)
						{
							1 {$xcol.width = 100; Break}
							2 {$xcol.width = 100; Break}
							3 {$xcol.width = 108; Break}
							4 {$xcol.width = 66; Break}
							5 {$xcol.width = 56; Break}
							6 {$xcol.width = 56; Break}
						}
					}

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
					$Table.AutoFitBehavior($wdAutoFitFixed)

					#Return focus back to document
					$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
					$TableRange = $Null
					$Table = $Null
				}
				If($Text)
				{
					Line 0 ""
				}
				If($HTML)
				{
					$columnWidths  = @( '125px', '125px', '135px', '70px', '60px', '60px' )
					$columnHeaders = @(
						'Name',                   $htmlsb,
						'SamAccountName',         $htmlsb,
						'Domain',                 $htmlsb,
						'Password Last Changed',  $htmlsb,
						'Password Never Expires', $htmlsb,
						'Account Enabled',        $htmlsb
					)

					FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '575'
					WriteHTMLLine 0 0 ''

					$rowData = $null
				}
			}
			ElseIf(!$?)
			{
				$txt = "Unable to retrieve users with AdminCount=1"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt "" $Null 0 $False $True
				}
				If($Text)
				{
					Line 0 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}
			}
			Else
			{
				$txt1 = "Users with AdminCount=1: "
				$txt2 = "<None>"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 4 0 $txt1
					WriteWordLIne 0 0 $txt2
				}
				If($Text)
				{
					Line 0 $txt1
					Line 0 $txt2
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 $txt1
					WriteHTMLLIne 0 0 "None"
				}
			}
			
			Write-Verbose "$(Get-Date -Format G): `t`tListing groups with AdminCount=1"
			#$AdminCounts = @(Get-ADGroup -LDAPFilter "(admincount=1)" -Server $Domain -EA 0 | Select-Object Name)
			#Fix in 3.03 where the group name doesn’t match the samAccountName or the distinguishedName
			$AdminCounts = @(Get-ADGroup -LDAPFilter "(admincount=1)" -Server $Domain -EA 0 | Select-Object Name, distinguishedName)
			
			If($? -and $Null -ne $AdminCounts)
			{
				$AdminCounts = $AdminCounts | Sort-Object Name
				[int]$AdminsCount = $AdminCounts.Count
				[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
				
				If($MSWORD -or $PDF)
				{
					WriteWordLine 4 0 "Groups with AdminCount=1 ($($AdminsCountStr) members):"
					$TableRange = $Script:doc.Application.Selection.Range
					[int]$Columns = 2
					[int]$Rows = $AdminCounts.Count + 1
					[int]$xRow = 1
					$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.AutoFitBehavior($wdAutoFitFixed)
					$Table.Style = $Script:MyHash.Word_TableGrid
			
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Group Name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Members"
				}
				If($Text)
				{
					Line 1 "Groups with AdminCount=1 ($($AdminsCountStr) members):"
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Groups with AdminCount=1 ($($AdminsCountStr) members):"
					#V3.00 FIXME - canNOT pre-allocate rowdata
					$rowdata = @()
				}
				ForEach($Admin in $AdminCounts)
				{
					Write-Verbose "$(Get-Date -Format G): `t`t`t$($Admin.Name)"
					If($MSWord -or $PDF)
					{
						$xRow++
					}
					
					#Fix in 3.03 where the group name doesn’t match the samAccountName or the distinguishedName
					#[array]$Members = @(Get-ADGroupMember -Identity $Admin.Name -Server $Domain -EA 0 | Sort-Object Name)
					
					try
					{
						[array]$Members = @(Get-ADGroupMember -Identity $Admin.distinguishedName -Server $Domain -EA 0 | Sort-Object Name)
						$GroupError = $False
					}
					
					catch
					{
						#hide error from console
						$GroupError = $True
					}
					
					If($? -and $Null -ne $Members)
					{
						[int]$MembersCount = $Members.Count
					}
					Else
					{
						[int]$MembersCount = 0
					}

					[string]$MembersCountStr = "{0:N0}" -f $MembersCount
					
					If($MSWord -or $PDF)
					{
						$Table.Cell($xRow,1).Range.Text = "$($Admin.Name) ($($MembersCountStr) members)"
						If($GroupError)
						{
							$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
							$Table.Cell($xRow,2).Range.Font.Bold  = $True
							$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
							$Table.Cell($xRow,2).Range.Text       = "Unable to retrieve group members. Check for orphaned SIDs."
						}
						Else
						{
							$MbrStr = ""
							If($MembersCount -gt 0)
							{
								ForEach($Member in $Members)
								{
									$MbrStr += "$($Member.Name)`r"
								}
								$Table.Cell($xRow,2).Range.Text = $MbrStr
							}
						}
					}
					If($Text)
					{
						[string]$MembersCountStr = "{0:N0}" -f $MembersCount
						Line 2 "Group Name`t: $($Admin.Name) ($($MembersCountStr) members)"
						If($GroupError)
						{
							Line 2 "Members`t`t: ***Unable to retrieve group members. Check for orphaned SIDs.***"
						}
						Else
						{
							$MbrStr = ""
							If($MembersCount -gt 0)
							{
								Line 2 "Members`t`t: " -NoNewLine
								$cnt = 0
								ForEach($Member in $Members)
								{
									$cnt++
									
									If($cnt -eq 1)
									{
										Line 0 $Member.Name
									}
									Else
									{
										Line 4 "  $($Member.Name)"
									}
								}
							}
							Else
							{
								Line 2 "Members`t`t: None"
							}
						}
						Line 0 ""
					}
					If($HTML)
					{
						[string]$MembersCountStr = "{0:N0}" -f $MembersCount
						$GroupName = "$($Admin.Name) ($($MembersCountStr) members)"
						If($GroupError)
						{
							$MbrStr = "Unable to retrieve group members. Check for orphaned SIDs."
							$rowdata += @(, (
								$GroupName, $htmlwhite,
								$MbrStr, $htmlred
							) )
						}
						Else
						{
							If($MembersCount -gt 0)
							{
								$first = $GroupName
								ForEach($Member in $Members)
								{
									$rowdata += @(, (
										$first,       $htmlwhite,
										$Member.Name, $htmlwhite
									) )
									$first = ''
								}
							}
							Else
							{
								$rowdata += @(, (
									$GroupName, $htmlwhite,
									'empty',  $htmlwhite
								) )
							}
						}
					}
				}
				
				If($MSWord -or $PDF)
				{
					#set column widths
					$xcols = $table.columns

					ForEach($xcol in $xcols)
					{
						Switch ($xcol.Index)
						{
						  1 {$xcol.width = 250; Break}
						  2 {$xcol.width = 175; Break}
						}
					}
					
					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
					$Table.AutoFitBehavior($wdAutoFitFixed)

					#Return focus back to document
					$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
					$TableRange = $Null
					$Table = $Null
				}
				If($Text)
				{
				}
				If($HTML)
				{
					$columnWidths  = @( '300px', '175px' )
					$columnHeaders = @(
						'Group Name', $htmlsb,
						'Members',    $htmlsb
					)

					FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '475'
					WriteHTMLLine 0 0 ''

					$rowData = $null
				}
			}
			ElseIf(!$?)
			{
				$txt = "Unable to retrieve Groups with AdminCount=1"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt "" $Null 0 $False $True
				}
				If($Text)
				{
					Line 0 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}
			}
			Else
			{
				$txt1 = 'Groups with AdminCount=1:'
				$txt2 = '<None>'
				If($MSWORD -or $PDF)
				{
					WriteWordLine 4 0 $txt1
					WriteWordLine 0 0 $txt2
				}
				If($Text)
				{
					Line 0 $txt1
					Line 0 $txt2
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 $txt1
					WriteHTMLLine 0 0 'None'
				}
			}
		}
		$First = $False
	}
} ## end Function ProcessGroupInformation
#endregion

#region GPOs by domain
Function ProcessGPOsByDomain
{
	Write-Verbose "$(Get-Date -Format G): Writing domain group policy data"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Group Policies by Domain"
	}
	If($Text)
	{
		Line 0 "///  Group Policies by Domain  \\\"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Group Policies by Domain&nbsp;&nbsp;\\\"
	}
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing group policies for domain $($Domain)"

		$DomainInfo = Get-ADDomain -Identity $Domain -EA 0 

		If( !$? )
		{
			$txt = "Error retrieving domain data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt '' $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If( $null -eq $DomainInfo )
		{
			$txt = "No Domain data was retrieved for domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}

		$txt = $Domain
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt += ' (Forest Root)'
		}

		If($MSWORD -or $PDF)
		{
			WriteWordLine 2 0 $txt
			WriteWordLine 3 0 "Linked Group Policy Objects" 
		}
		If($Text)
		{
			Line 1 "///  $($txt)  \\\"
			Line 0 "Linked Group Policy Objects" 
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
			WriteHTMLLine 3 0 "Linked Group Policy Objects" 
		}

		Write-Verbose "$(Get-Date -Format G): `t`tGetting linked GPOs"

		$LinkedGPOs = @($DomainInfo.LinkedGroupPolicyObjects | Sort-Object)
		If($Null -eq $LinkedGpos)
		{
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 "<None>"
			}
			If($Text)
			{
				Line 2 "<None>"
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "None"
			}
		}
		Else
		{
			#V3.00 pre-allocate GPOArray
			## $GPOArray = New-Object System.Collections.ArrayList
			$GPOArray = New-Object System.Array[] $LinkedGpos.Count
			$gpoIndex = 0

			ForEach($LinkedGpo in $LinkedGpos)
			{
				#taken from Michael B. Smith's work on the XenApp 6.x scripts
				#this way we don't need the GroupPolicy module
				$gpObject = [ADSI]( "LDAP://" + $LinkedGPO )
				If($Null -eq $gpObject.DisplayName)
				{
					$p1 = $LinkedGPO.IndexOf("{")
					#38 is length of guid (32) plus the four "-"s plus the beginning "{" plus the ending "}"
					$GUID = $LinkedGPO.SubString($p1,38)
					$tmp = "GPO with GUID $($GUID) was not found in this domain"
				}
				Else
				{
					$tmp = $gpObject.DisplayName	### name of the group policy object
				}
				#V3.00
				## $GPOArray.Add($tmp) > $Null
				$GPOArray[ $gpoIndex ] = $tmp
				$gpoIndex++
				$gpObject = $null
			}

			$GPOArray = $GPOArray | Sort-Object 

			If($MSWORD -or $PDF)
			{
				$ItemsWordTable = New-Object System.Collections.ArrayList
				ForEach($Item in $GPOArray)
				{
					## Add the required key/values to the hashtable
					$WordTableRowHash = @{ 
						GPOName = $Item
					}

					## Add the hash to the array
					$ItemsWordTable.Add($WordTableRowHash) > $Null
				}

				If($ItemsWordTable.Count -gt 0)	#3.07
				{
					$Table = AddWordTable -Hashtable $ItemsWordTable `
					-Columns GPOName `
					-Headers "GPO Name" `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed

					SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 300

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
				}
			}
			If($Text)
			{
				ForEach($Item in $GPOArray)
				{
					Line 2 $Item
				}
				Line 0 ""
			}
			If($HTML)
			{
				#V3.00 - pre-allocate rowdata
				## $rowdata = @()
				$rowData  = New-Object System.Array[] $GPOArray.Count
				$rowIndx = 0

				ForEach($Item in $GPOArray)
				{
					$rowdata[ $rowIndx ] = @( $Item, $htmlwhite )
					$rowIndx++
				}

				$columnHeaders = @( 'GPO Name', $htmlsb )

				FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders
				WriteHTMLLine 0 0 ''

				$rowData = $null
			}
			$GPOArray = $Null
		}
		$LinkedGPOs = $Null
		$First = $False
	}
} ## end Function ProcessGPOsByDomain
#endregion

#region group policies by organizational units
Function ProcessgGPOsByOUOld
{
	Write-Verbose "$(Get-Date -Format G): Writing Group Policy data by Domain by OU"
	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Group Policies by Organizational Unit"
	}
	If($Text)
	{
		Line 0 "///  Group Policies by Organizational Unit  \\\"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Group Policies by Organizational Unit&nbsp;&nbsp;\\\"
	}
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing domain $($Domain)"
		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}

		$Disclaimer = "(Contains only OUs with linked Group Policies)"
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt = "Group Policies by OUs in Domain $($Domain) (Forest Root)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
				#print disclaimer line in 8 point bold italics
				WriteWordLine 0 0 $Disclaimer "" $Null 8 $True $True
			}
			If($Text)
			{
				Line 0 "///  $($txt)  \\\"
				Line 0 $Disclaimer
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
				WriteHTMLLine 0 0 $Disclaimer -option $htmlBold
			}
		}
		Else
		{
			$txt = "Group Policies by OUs in Domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
				WriteWordLine 0 0 $Disclaimer "" $Null 8 $True $True
			}
			If($Text)
			{
				Line 1 "///  $($txt)  \\\"
				Line 1 $Disclaimer
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
				WriteHTMLLine 0 0 $Disclaimer -options $htmlBold
			}
		}
		
		#V3.00
		Write-Verbose "$(Get-Date -Format G): `tSearching for all OUs in domain $($Domain)"

		## FIXME - we get "all OUs for the domain" three times in this script - that needs to be fixed.
		## [Webster] not really. ProcessGPOsByOUNew and ProcessGPOsByOUOld are separate Functions and never used in the same script run
		## ProcessOrganizationUnits uses different OU properties. So, $OUs is only used twice but each is different.
		
		## FIXME - v3.00 see optimizations applied in getDSUsers
		#get all OUs for the domain
		$OUs = @(Get-ADOrganizationalUnit -Filter * -Server $Domain `
			-Properties CanonicalName, DistinguishedName, Name -EA 0 | `
			Select-Object CanonicalName, DistinguishedName, Name | `
			Sort-Object CanonicalName)
		
		If( !$? )
		{
			Write-Warning "Error retrieving OU base data for OU $($OU.CanonicalName)"
			Continue
		}

		If( $null -eq $OUs )
		{
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 "<None>"
			}
			If($Text)
			{
				Line 0 "<None>"
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "None"
			}

			Continue
		}

		[int]$NumOUs = $OUs.Count
		[int]$OUCount = 0

		#V3.03
		Write-Verbose "$(Get-Date -Format G): `tThere are $NumOUs OUs in domain $($Domain)"

		ForEach($OU in $OUs)
		{
			$OUCount++
			$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing OU $($OU.CanonicalName) - OU # $OUCount of $NumOUs"
			
			#V3.00 FIXME MAYBE???? LinkedGroupPolicyObjects
			#get data for the individual OU
			##$OUInfo = Get-ADOrganizationalUnit -Identity $OU.DistinguishedName -Server $Domain -Properties * -EA 0 
			$OUInfo = Get-ADOrganizationalUnit -Identity $OU.DistinguishedName -Server $Domain `
				-Properties LinkedGroupPolicyObjects -EA 0 
			
			If( !$? )
			{
				$txt = "Error retrieving OU GPO data for domain $($Domain)"
				Write-Warning $txt
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt "" $Null 0 $False $True
				}
				If($Text)
				{
					Line 0 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}

				Continue
			}

			If( $null -eq $OUInfo )
			{
				$txt = "No OU data was retrieved for domain $($Domain)"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt "" $Null 0 $False $True
				}
				If($Text)
				{
					Line 0 $txt
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}

				Continue
			}
	
			Write-Verbose "$(Get-Date -Format G): `t`t`tGetting linked GPOs"
			[array]$LinkedGPOs = $OUInfo.LinkedGroupPolicyObjects
			If($LinkedGpos.Count -eq 0)
			{
				# do nothing
			}
			Else
			{
				$LinkedGPOs = $LinkedGPOs | Sort-Object
				#V3.00 use pre-allocated GPOArray
				## $GPOArray = New-Object System.Collections.ArrayList
				$GPOArray = New-Object System.Array[] $LinkedGpos.Count
				$gpoIndex = 0

				ForEach($LinkedGpo in $LinkedGpos)
				{
					#taken from Michael B. Smith's work on the XenApp 6.x scripts
					#this way we don't need the GroupPolicy module
					$gpObject = [ADSI]( "LDAP://" + $LinkedGPO )
					If($Null -eq $gpObject.DisplayName)
					{
						$p1 = $LinkedGPO.IndexOf("{")
						#38 is length of guid (32) plus the four "-"s plus the beginning "{" plus the ending "}"
						$GUID = $LinkedGPO.SubString($p1,38)
						$tmp = "GPO with GUID $($GUID) was not found in this domain"
					}
					Else
					{
						$tmp = $gpObject.DisplayName	### name of the group policy object
					}
					#V3.00
					## $GPOArray.Add($tmp) > $Null
					$GPOArray[ $gpoIndex ] = $tmp
					$gpoIndex++
					$gpObject = $null
				}

				$GPOArray = $GPOArray | Sort-Object 

				[int]$Rows = $LinkedGpos.Count
				
				If($MSWORD -or $PDF)
				{
					[int]$Columns = 1

					WriteWordLine 3 0 "$($OUDisplayName) ($($Rows))"
					$TableRange = $Script:doc.Application.Selection.Range
					$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.Style = $Script:MyHash.Word_TableGrid
	
					$Table.Borders.InsideLineStyle = $wdLineStyleNone
					$Table.Borders.OutsideLineStyle = $wdLineStyleNone
					
					[int]$xRow = 0
					ForEach($Item in $GPOArray)
					{
						$xRow++
						#3.07 somehow PoSH thinks some GPOs are arrays. 
						#The GPOs in my lab that start with "+" PoSH says are arrays and I started getting this error
						#Unable to cast object of type 'System.Object[]' to type 'System.String'.
						#At C:\Webster\ADDS_Inventory_V3.ps1:14923 char:7
						#+                         $Table.Cell($xRow,1).Range.Text = $Item
						#+                         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						#+ CategoryInfo          : OperationStopped: (:) [], InvalidCastException
						#+ FullyQualifiedErrorId : System.InvalidCastException       
						#
						#Group Policies by OUs in Domain LabADDomain.com (Forest Root)
						#(Contains only OUs with linked Group Policies)
						#Domain Controllers (3)
						#+ SERVER Set PDCe Domain Controller as Authoritative Time Server v1.0 [somehow seen as an array, not a string]
						#+ SERVER Set Time Settings on non-PDCe Domain Controllers v1.0 [somehow seen as an array, not a string]
						#Default Domain Controllers Policy

                        If($Item -is [array]) 
                        {
						    $Table.Cell($xRow,1).Range.Text = $Item[0]
                        }
                        Else
                        {
						    $Table.Cell($xRow,1).Range.Text = $Item
                        }
					}

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
					$Table.AutoFitBehavior($wdAutoFitContent)

					#Return focus back to document
					$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
					$TableRange = $Null
					$Table = $Null
				}
				If($Text)
				{
					Line 2 "$($OUDisplayName) ($($Rows))"
					ForEach($Item in $GPOArray)
					{
						Line 3 $Item
					}
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "$($OUDisplayName) ($($Rows))"
					#V3.00 - pre-allocate rowdata
					## $rowdata = @()
					#$rowData = New-Object System.Array[] $GPOArray.Count
					#$rowIndx = 0

					#ForEach($Item in $GPOArray)
					#{
					#	$rowdata[ $rowIndx ] = @( $Item, $htmlwhite )
					#	$rowIndx++
					#}

					#$columnHeaders = @(
					#	'GPO Name', $htmlsb
					#)

					#FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders
					$rowdata = @()
					ForEach($Item in $GPOArray)
					{
						$rowdata += @(,($Item,$htmlwhite))
					}
					$columnHeaders = @('GPO Name',($htmlsilver -bor $htmlbold))
					$msg = ""
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 ''

					$rowData = $null
				}
				$GPOArray = $null
			}
			$LinkedGPOs = $null
		}
		$First = $False
		$OUs   = $null
	}
} ## end Function ProcessgGPOsByOUOld

Function ProcessgGPOsByOUNew
{
	Write-Verbose "$(Get-Date -Format G): Writing Group Policy data by Domain by OU"
	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Group Policies by Organizational Unit"
	}
	If($Text)
	{
		Line 0 "///  Group Policies by Organizational Unit  \\\"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Group Policies by Organizational Unit&nbsp;&nbsp;\\\"
	}
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing domain $($Domain)"
		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}

		$Disclaimer = "(Contains only OUs with linked or inherited Group Policies)"
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt = "Group Policies by OUs in Domain $($Domain) (Forest Root)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
				#print disclaimer line in 8 point bold italics
				WriteWordLine 0 0 $Disclaimer "" $Null 8 $True $True
			}
			If($Text)
			{
				Line 0 "///  $($txt)  \\\"
				Line 0 $Disclaimer
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
				WriteHTMLLine 0 0 $Disclaimer -options $htmlBold
			}
		}
		Else
		{
			$txt = "Group Policies by OUs in Domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
				WriteWordLine 0 0 $Disclaimer "" $Null 8 $True $True
			}
			If($Text)
			{
				Line 1 "///  $($txt)  \\\"
				Line 1 $Disclaimer
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
				WriteHTMLLine 0 0 $Disclaimer -options $htmlBold
			}
		}

		#V3.00
		Write-Verbose "$(Get-Date -Format G): `tSearching for all OUs in domain $($Domain)"

		#get all OUs for the domain
		$OUs = @(Get-ADOrganizationalUnit -Filter * -Server $Domain `
			-Properties CanonicalName, DistinguishedName, Name -EA 0 | `
			Select-Object CanonicalName, DistinguishedName, Name | `
			Sort-Object CanonicalName)

		If( !$? )
		{
			$txt = "Error retrieving OU data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If( $null -eq $OUs )
		{
			$txt = "No OU data was retrieved for domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		[int]$NumOUs = $OUs.Count
		[int]$OUCount = 0

		#V3.00
		Write-Verbose "$(Get-Date -Format G): `tThere are $NumOUs OUs in domain $($Domain)"

		ForEach($OU in $OUs)
		{
			$OUCount++
			$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing OU $($OU.CanonicalName) - OU # $OUCount of $NumOUs"
			Write-Verbose "$(Get-Date -Format G): `t`t`tGetting linked and inherited GPOs"

			#change for 2.16
			#work around invalid property DisplayName when the gpolinks and inheritedgpolinks collections are empty

			$Results = Get-GPInheritance -target $OU.DistinguishedName -Domain $Domain -EA 0 #3.04 add domain

			#V3.00 check for error Return, just like with Get-AdOrganizationalUnit above
			If( !$? )
			{
				Write-Warning "Error retrieving OU GPO data for OU $($OU.CanonicalName)"
				Continue
			}

			If( $null -eq $Results )
			{
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 "<None>"
				}
				If($Text)
				{
					Line 0 "<None>"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "None"
				}

				Continue
			}

			$LinkedGPOs = $Null
			If(($Results.GpoLinks).Count -gt 0)
			{
				$LinkedGPOs = $Results.GpoLinks.DisplayName  ## depends on automated unrolling - FIXME
			}

			$InheritedGPOs = $Null
			If(($Results.InheritedGpoLinks).Count -gt 0)
			{
				$InheritedGPOs = $Results.InheritedGpoLinks.DisplayName  ## depends on automated unrolling - FIXME
			}

			If($Null -eq $LinkedGPOs -and $Null -eq $InheritedGPOs)
			{
				# do nothing
			}
			Else
			{
				#V3.00 Switch to pre-allocated
				## $AllGPOs  = New-Object System.Collections.ArrayList[] $InheritedGPOs.Length
				## wv "***** ProcessGPOsByOUNew InheritedGPOs.Length $( $InheritedGPOs.Length )"
				$AllGPOs = New-Object System.Array[] $InheritedGPOs.Length
				$gpoIndex = 0

				ForEach($item in $InheritedGPOs)
				{
					## $obj = New-Object -TypeName PSObject
					$GPOType = ''
					If(!($LinkedGPOs -contains $item))
					{
						$GPOType = 'Inherited'
					}
					Else
					{
						$GPOType = 'Linked'
					}
					$AllGPOs[ $gpoIndex ] = [PSCustomObject] @{ 
						GPOName = $item
						GPOType = $GPOType
					}
					$gpoIndex++
				}

				$AllGPOS = $AllGPOs | Sort-Object GPOName

				If($null -eq $AllGPOS)
				{
					[int]$Rows = 0
				}
				Else
				{
					[int]$Rows = $AllGPOS.Length
				}

				If($MSWORD -or $PDF)
				{
					WriteWordLine 3 0 "$($OUDisplayName) ($($Rows))"
					$GPOWordTable = New-Object System.Collections.ArrayList
					ForEach($Item in $AllGPOS)
					{
						## Add the required key/values to the hashtable
						$WordTableRowHash = @{ 
						GPOName = $Item.GPOName; 
						GPOType = $Item.GPOType 
						}

						## Add the hash to the array
						$GPOWordTable.Add($WordTableRowHash) > $Null
					}

					If($GPOWordTable.Count -gt 0)	#3.07
					{
						$Table = AddWordTable -Hashtable $GPOWordTable `
						-Columns GPOName, GPOType `
						-Headers "GPO Name", "GPO Type" `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						## IB - set column widths without recursion
						$Table.Columns.Item(1).Width = 400;
						$Table.Columns.Item(2).Width = 65;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
				}
				If($Text)
				{
					Line 2 "$($OUDisplayName) ($($Rows))"
					ForEach($Item in $AllGPOS)
					{
						Line 3 "Name: " $Item.GPOName
						Line 3 "Type: " $Item.GPOType
						Line 0 ""
					}
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "$($OUDisplayName) ($($Rows))"
					#V3.00 Switch to pre-allocate rowdata
					## $rowdata = @()
					$rowData = New-Object System.Array[] $AllGPOs.Length
					$rowIndx = 0

					ForEach($Item in $AllGPOS)
					{
						$rowdata[ $rowIndx ] = @(
							$Item.GPOName, $htmlwhite,
							$Item.GPOType, $htmlwhite
						)
						$rowIndx++
					}

					$columnWidths  = @( '700px', '65px' )
					$columnHeaders = @(
						'GPO Name', $htmlsb,
						'GPO Type', $htmlsb
					)

					FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '765'
					WriteHTMLLine 0 0 ''

					$rowData = $null
				}
				$AllGPOS = $null
			}
		}
		$OUs   = $null
		$First = $False
	}
} ## end Function ProcessgGPOsByOUNew
#endregion

#region OUs by domain for Blocked Inheritance
Function ProcessOUsForBlockedInheritance
{
	Write-Verbose "$(Get-Date -Format G): Writing OUs with GPO Blocked Inheritance"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "OUs with GPO Blocked Inheritance"
	}
	If($Text)
	{
		Line 0 "///  OUs with GPO Blocked Inheritance  \\\"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;OUs with GPO Blocked Inheritance&nbsp;&nbsp;\\\"
	}
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing OUs with GPO Blocked Inheritance for domain $($Domain)"

		$DomainInfo = Get-ADDomain -Identity $Domain -EA 0 

		If( !$? )
		{
			$txt = "Error retrieving domain data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt '' $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If( $null -eq $DomainInfo )
		{
			$txt = "No Domain data was retrieved for domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}

		$txt = $Domain
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt += ' (Forest Root)'
		}

		If($MSWORD -or $PDF)
		{
			WriteWordLine 2 0 $txt
			WriteWordLine 3 0 "OUs with Blocked Inheritance" 
		}
		If($Text)
		{
			Line 1 "///  $($txt)  \\\"
			Line 1 "OUs with Blocked Inheritance" 
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
			WriteHTMLLine 3 0 "OUs with Blocked Inheritance" 
		}

		Write-Verbose "$(Get-Date -Format G): `t`tGetting OUs with Blocked Inheritance"

		$DomainName = $DomainInfo.Name
		#$DomainFQDN = $DomainInfo.DNSRoot
		
		$source             = New-Object System.DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
		$source.SearchScope = 'Subtree'
		$source.PageSize    = 1000
		$source.filter      = '(&(objectclass=OrganizationalUnit)(gpoptions=1))'

		$BlockedOUs = $source.findall()
		
		If($Null -eq $BlockedOUs)
		{
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 "<None>"
			}
			If($Text)
			{
				Line 2 "<None>"
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "None"
			}
		}
		Else
		{
			$OUArray = New-Object System.Collections.ArrayList
			
			ForEach($Item in $BlockedOUs)
			{
				$obj = [PSCustomObject] @{
					OUName = $Item.Properties.Item('distinguishedname')
				}

				$null = $OUArray.Add( $obj )
			}
			
			$OUArray = @($OUArray | Sort-Object OUName)
			
			If($MSWORD -or $PDF)
			{
				$ItemsWordTable = New-Object System.Collections.ArrayList
				ForEach($Item in $OUArray)
				{
					## Add the required key/values to the hashtable
					$WordTableRowHash = @{ 
						OUName = $Item.OUName
					}

					## Add the hash to the array
					$ItemsWordTable.Add($WordTableRowHash) > $Null
				}

				If($ItemsWordTable.Count -gt 0)	#3.07
				{
					$Table = AddWordTable -Hashtable $ItemsWordTable `
					-Columns OUName `
					-Headers "OU Name" `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed

					SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 300

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
				}
			}
			If($Text)
			{
				ForEach($Item in $OUArray)
				{
					Line 2 $Item.OUName
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowData = @()

				ForEach($Item in $OUArray)
				{
					$rowdata += @(,($Item.OUName,$htmlwhite))
				}

				$columnHeaders = @( 'OU Name', $htmlsb )

				FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders
				WriteHTMLLine 0 0 ''

				$rowData = $null
			}

			$OUArray = $Null
		}
		$First = $False
	}
} ## end Function ProcessGPOsByDomain
#endregion

#region misc info by domain
## added for v3.00
$TsHomeDrive  = 'TerminalServicesHomeDrive'
$TsHomeDir    = 'TerminalServicesHomeDirectory'
$TsProfPath   = 'TerminalServicesProfilePath'
$TsAllowLogon = 'AllowLogon'

## added for v3.00
Function GetTsAttributes
{
	Param
	(
		[Parameter( Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true )]
		[string] $distinguishedName
	)

	Process
	{
		$u = [ADSI] ( 'LDAP://' + $distinguishedName )
		If( $u )
		{
			If( $u.psbase.InvokeGet( 'userParameters' ) )
			{
				###FIXME: Need to add error checking or validation of user accounts here. A corrupt user account or
				###		  a user account with corrupt TsAttributes causes an internal PoSH error not seen in the console but
				###		  using -dev records the error
				$o = @{
					$TsHomeDrive      = $u.psbase.InvokeGet( $TsHomeDrive )
					$TsHomeDir        = $u.psbase.InvokeGet( $TsHomeDir )
					$TsProfPath       = $u.psbase.InvokeGet( $TsProfPath )
					$TsAllowLogon     = $u.psbase.InvokeGet( $TsAllowLogon ) -As [Bool]
					DistinguishedName = $u.distinguishedname.value
					SamAccountName    = $u.samaccountname.value
				}
			}
			Else
			{
				$o = @{
					$TsHomeDrive      = $null
					$TsHomeDir        = $null
					$TsProfPath       = $null
					$TsAllowLogon     = $null
					DistinguishedName = $u.distinguishedname.value
					SamAccountName    = $u.samaccountname.value
				}
			}

			Write-Output $o
			$o = $null
			$u = $null
		}
	}
} ## end Function GetTsAttributes

Function getDSUsers
{
    [CmdletBinding()]
    Param
    (
        [String] $TrustedDomain
    )

	[Int64] $script:MaxPasswordAge = $null


	[Object] $domainADSI = $null

	Function GetMaximumPasswordAge
	{
		###
		### GetMaximumPasswordAge
		###
		### Retrieve the maximum password age that is set on the domain object. This is
		### normally set by the "Default Domain Policy".
		###

		If( $null -ne $MaxPasswordAge -and $MaxPasswordAge -gt 0 )
		{
			### Cache the value so that it only has to be retrieved once, converted
			### to an int64 once, and converted to days once. Win-win-win.

			Return $MaxPasswordAge
		}

		### Dealing with ADSI unfortunately also means dealing with COM objects.
		### Using ConvertLargeIntegerToInt64 takes the COM object and converts 
		### it into a native .Net type.

		#Uncomment this line in 3.11 to fix bug reported by Danny de Kooker
		[Int64] $script:MaxPasswordAge = $domainADSI.ConvertLargeIntegerToInt64( $domainADSI.maxPwdAge.Value )

		### Convert to days
		### 	there are 86,400 seconds per day (24 * 60 * 60)
		### there are 10,000,00 100 nanosecond units per second
		### https://docs.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-filetime

		#Update for 3.03 from MBS

		#Comment this line in 3.11 to fix bug reported by Danny de Kooker
		#[Int64] $script:MaxPasswordAge = ( -$MaxPasswordAge / ( 86400 * 10000000 ) )
		
		If( $script:MaxPasswordAge -eq [Int64]::MinValue )
		{
			### there is no defined maximum password age group policy
			$script:MaxPasswordAge = -1
		}
		Else
		{
			### Convert to days
			###     there are 86,400 seconds per day (24 * 60 * 60)
			###     there are 10,000,000 100 nanosecond units per second
			### https://docs.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-filetime

			[Int64] $script:MaxPasswordAge = ( -$MaxPasswordAge / ( 86400 * 10000000 ) )
		}
		Return $MaxPasswordAge
	}

    #$script_begin = Get-Date

    ## so this block of code takes a FQDN and turns it into a distinguishedName
    ## e.g., fabrikam.com --> DC=fabrikam,DC=com
    ## FIXME MBS - when I have a minute - this is serious overkill
    $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext( 'Domain', $TrustedDomain )
    try 
    {
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain( $context )
    }
    catch [Exception] 
    {
        Write-Error $_.Exception.Message
        AbortScript
    }

    # Get AD Distinguished Name
    $ADSearchBase = $Domain.GetDirectoryEntry().DistinguishedName.Value
	$domainADSI   = [ADSI]( 'LDAP://' + $ADSearchBase )
	$null         = GetMaximumPasswordAge
	$now          = Get-Date

    Write-Verbose "$(Get-Date -Format G): `t`tGathering user misc data #1"

    ## see MBS Get-myUserInfo.ps1 for the full list
    $ADS_UF_ACCOUNTDISABLE                          = 2        ### 0x2
    $ADS_UF_LOCKOUT                                 = 16       ### 0x10
    $ADS_UF_PASSWD_NOTREQD                          = 32       ### 0x20
    $ADS_UF_PASSWD_CANT_CHANGE                      = 64       ### 0x40
##  $ADS_UF_NORMAL_ACCOUNT                          = 512      ### 0x200
    $ADS_UF_DONT_EXPIRE_PASSWD                      = 65536    ### 0x10000
    $ADS_UF_PASSWORD_EXPIRED                        = 8388608  ### 0x800000

    $WellKnownPrimaryGroupIDs = 
    @{
        512 = 'Domain Admins'
        513 = 'Domain Users'
        514 = 'Domain Guests'
        515 = 'Domain Computers'
        516 = 'Domain Controllers'
        517 = 'Cert Publishers'
        518 = 'Schema Admins'
        519 = 'Enterprise Admins'
    }

    $ADPropertyList = 
    @(
        'distinguishedname'
        'samaccountName'
        'userprincipalname'
        'lastlogontimestamp'
        'homedrive'
        'homedirectory'
        'profilepath'
        'scriptpath'
        'primarygroupid'        ## only the RID of the primary group
        'useraccountcontrol'
        'sidhistory'
        'pwdlastset'
        'userparameters'
    )

    $ADSearchRoot           = New-Object System.DirectoryServices.DirectoryEntry( 'LDAP://' + $ADSearchBase ) 
    $ADSearcher             = New-Object System.DirectoryServices.DirectorySearcher
    $ADSearcher.SearchRoot  = $ADSearchRoot
    $ADSearcher.PageSize    = 1000
    $ADSearcher.Filter      = '(&(objectcategory=person)(objectclass=user))' ## all user objects
    $ADSearcher.SearchScope = 'subtree'

    If( $ADPropertyList ) 
    {
        ForEach( $ADProperty in $ADPropertyList ) 
        {
            $null = $ADSearcher.PropertiesToLoad.Add( $ADProperty )
        }
    }

    try 
    {
        Write-Verbose "Please be patient whilst the script retrieves all user objects and specified attributes..."
        $colResults = $ADSearcher.Findall()
        # Dispose of the search and results properly to avoid a memory leak
        $ADSearcher.Dispose()
        $ctUsers = $colResults.Count
    }
    catch 
    {
        $ctUsers = 0
        Write-Warning "search failed, $( $error[ 0 ] )"
    }

    If( $ctUsers -eq 0 )
    {
        Write-Verbose "$(Get-Date -Format G): `t`tNo users found, exiting"

        Return
    }

    If( $ctUsers -gt 50000 )
    {
        Write-Verbose "$(Get-Date -Format G): `t`t`t******************************************************************************************************"
        Write-Verbose "$(Get-Date -Format G): `t`t`tThere are $ctUsers user accounts to process. Building user lists will take a long time. Be patient."
        Write-Verbose "$(Get-Date -Format G): `t`t`t******************************************************************************************************"
    }
    Else
    {
        Write-Verbose "$(Get-Date -Format G): Processing $ctUsers user objects in the $domain Domain..."        
    }
    
    $ctUsersDisabled              = 0
    $ctUsersUnknown               = 0
    $ctUsersLockedOut             = 0
    $ctPasswordExpired            = 0
    $ctPasswordNeverExpires       = 0
    $ctPasswordNotRequired        = 0
    $ctCannotChangePassword       = 0
    $ctNolastLogonTimestamp       = 0
    $ctHasSIDHistory              = 0
    $ctHomeDrive                  = 0
    $ctPrimaryGroup               = 0
    $ctRDSHomeDrive               = 0
	$ctOrphanedFSPs               = 0

    $ctActiveUsers                = 0
    $ctActivePasswordExpired      = 0
    $ctActivePasswordNeverExpires = 0
    $ctActivePasswordNotRequired  = 0
    $ctActiveCannotChangePassword = 0
    $ctActiveNolastLogonTimestamp = 0

    $listUser                 = New-Object System.Collections.Generic.List[PsCustomObject] $ctUsers
    $listUsersDisabled        = New-Object System.Collections.Generic.List[PsCustomObject]
    $listUsersUnknown         = New-Object System.Collections.Generic.List[PsCustomObject]
    $listUsersLockedOut       = New-Object System.Collections.Generic.List[PsCustomObject]
    $listPasswordExpired      = New-Object System.Collections.Generic.List[PsCustomObject]
    $listPasswordNeverExpires = New-Object System.Collections.Generic.List[PsCustomObject]
    $listPasswordNotRequired  = New-Object System.Collections.Generic.List[PsCustomObject]
    $listCannotChangePassword = New-Object System.Collections.Generic.List[PsCustomObject]
    $listNolastLogonTimestamp = New-Object System.Collections.Generic.List[PsCustomObject]
    $listHasSIDHistory        = New-Object System.Collections.Generic.List[PsCustomObject]
    $listHomeDrive            = New-Object System.Collections.Generic.List[PsCustomObject]
    $listPrimaryGroup         = New-Object System.Collections.Generic.List[PsCustomObject]
    $listRDSHomeDrive         = New-Object System.Collections.Generic.List[PsCustomObject]
    $listOrphanedFSPs         = New-Object System.Collections.Generic.List[PsCustomObject]

##$global:colResults = $colresults

    $ctIndex = 0

    ForEach( $objResult in $colResults ) 
    {
        $ctIndex++
        If( ( $ctIndex % 5000 ) -eq 0 )
        {
            Write-Verbose "$(Get-Date -Format G): about to process user $ctIndex of $ctUsers"
        }

        $distinguishedname  = If( $null -ne ( $obj = $objResult.Properties[ 'distinguishedname' ] ) -and $obj.Count -gt 0 ) 
                            { $obj.Item( 0 ) } Else { $null }
        $useraccountcontrol = If( $null -ne ( $obj = $objResult.Properties[ 'useraccountcontrol' ] ) -and $obj.Count -gt 0 ) 
                            { $obj.Item( 0 ) } Else { $null }
        $samaccountname     = If( $null -ne ( $obj = $objResult.Properties[ 'samaccountname' ] ) -and $obj.Count -gt 0 ) 
                            { $obj.Item( 0 ) } Else { $null }

        $Unknown = $false
        If( $null -eq $useraccountcontrol )
        {
            $Unknown = $true
            $ctUsersUnknown++
        }

        If( $useraccountcontrol -band $ADS_UF_ACCOUNTDISABLE )
        {
            $Enabled = $false
            $ctUsersDisabled++
        }
        Else
        {
            $Enabled = $true
            $ctActiveUsers++
        }

        $passwordNeverExpires = $false
        If( $userAccountControl -band $ADS_UF_DONT_EXPIRE_PASSWD )
        {
            $passwordNeverExpires = $true
            $ctPasswordNeverExpires++
            If( $Enabled )
            {
                $ctActivePasswordNeverExpires++
            }
        }

        $passwordNotRequired = $false
        If( $userAccountControl -band $ADS_UF_PASSWD_NOTREQD )
        {
            $passwordNotRequired = $true
            $ctPasswordNotRequired++
            If( $Enabled )
            {
                $ctActivePasswordNotRequired++
            }
        }

        $LockedOut = $false
        If( $useraccountcontrol -band $ADS_UF_LOCKOUT )
        {
            $LockedOut = $true
            $ctUsersLockedOut++
        }

        $cannotChangePassword = $false
        If( $useraccountcontrol -band $ADS_UF_PASSWD_CANT_CHANGE )
        {
            $cannotChangePassword = $true
            $ctCannotChangePassword++
            If( $Enabled )
            {
                $ctActiveCannotChangePassword++
            }
        }

		$passwordExpired = $false
		If( $passwordNeverExpires -eq $false )
		{
			If( $useraccountcontrol -band $ADS_UF_PASSWORD_EXPIRED )
			{
				$passwordExpired = $true
			}
			ElseIf( $script:MaxPasswordAge -ge 0 )
			{
				$pls = If( $null -ne ( $obj = $objResult.Properties[ 'pwdlastset' ] ) -and $obj.Count -gt 0 ) 
					{ $obj.Item( 0 ) } Else { $null }
				$date = [DateTime] $pls
				$passwordLastSet = $date.AddYears( 1600 ).ToLocalTime()
				$passwordExpires = $passwordLastSet.AddDays( $script:MaxPasswordAge )
				If( $now -gt $passwordExpires )
				{
					$passwordExpired = $true
				}
				## write-verbose "***** sam $samaccountname, lastSet $($passwordLastSet), expires $($passwordExpires), expired $passwordexpired"
			}
			If( $passwordExpired )
			{
				$ctPasswordExpired++
				If( $Enabled )
				{
					$ctActivePasswordExpired++
				}
			}
		}

        $primaryGroupID = If( $null -ne ( $obj = $objResult.Properties[ 'primarygroupid' ] ) -and $obj.Count -gt 0 ) 
                        { $obj.Item( 0 ) } Else { $null }
        If( $WellKnownPrimaryGroupIDs.ContainsKey( $primaryGroupID ) )
        {
            $primaryGroup = $WellKnownPrimaryGroupIDs[ $primaryGroupID ]
        }
        Else
        {
            $primaryGroup = 'RID:' + $primaryGroupID.ToString()
        }
        If( $primaryGroupID -ne 513 )
        {
            $ctPrimaryGroup++
        }

        $lastlogontimestamp = If( $null -ne ( $obj = $objResult.Properties[ 'lastlogontimestamp' ] ) -and $obj.Count -gt 0 ) 
                            { $obj.Item( 0 ) } Else { $null }
        If( $null -eq $lastlogontimestamp )
        {
            $ctNolastLogonTimestamp++
            If( $Enabled )
            {
                $ctActiveNolastLogonTimestamp++
            }
        }

        $hasSIDHistory = If( $null -ne ( $obj = $objResult.Properties[ 'sidhistory' ] ) -and $obj.Count -gt 0 ) 
                        { $obj.Item( 0 ) } Else { $null }
        If( $null -ne $hasSIDHistory )
        {
            $ctHasSIDHistory++
        }

        $homedrive = If( $null -ne ( $obj = $objResult.Properties[ 'homedrive' ] ) -and $obj.Count -gt 0 ) 
                    { $obj.Item( 0 ) } Else { $null }
        If( $null -ne $homedrive )
        {
            $ctHomeDrive++
        }

        $homedirectory = If( $null -ne ( $obj = $objResult.Properties[ 'homedirectory' ] ) -and $obj.Count -gt 0 ) 
                        { $obj.Item( 0 ) } Else { $null }
        $profilepath   = If( $null -ne ( $obj = $objResult.Properties[ 'profilepath' ] ) -and $obj.Count -gt 0 ) 
                        { $obj.Item( 0 ) } Else { $null }
        $scriptpath    = If( $null -ne ( $obj = $objResult.Properties[ 'scriptpath' ] ) -and $obj.Count -gt 0 ) 
                        { $obj.Item( 0 ) } Else { $null }

        ## RDSHomeDrive is left
        $r_homedrive  = $null
        $r_homedir    = $null
        $r_profpath   = $null
        $r_allowlogon = $null

        $userparameters = If( $null -ne ( $obj = $objResult.Properties[ 'userparameters' ] ) -and $obj.Count -gt 0 ) 
                        { $obj.Item( 0 ) } Else { $null }
        If( $null -ne $userparameters )
        {
            ## TS/RDS/Citrix values are only present if the userparameters attribute exists
            $o = GetTsAttributes $distinguishedname
            If( $o )
            {
                $r_homedrive  = If( $o.ContainsKey( $TsHomeDrive  ) ) { $o[ $TsHomeDrive ]  } Else { $null }
                $r_homedir    = If( $o.ContainsKey( $TsHomeDir    ) ) { $o[ $TsHomeDir ]    } Else { $null }
                $r_profpath   = If( $o.ContainsKey( $TsProfPath   ) ) { $o[ $TsProfPath ]   } Else { $null }
                $r_allowlogon = If( $o.ContainsKey( $TsAllowLogon ) ) { $o[ $TsAllowLogon ] } Else { $null }
            }
        }

        If( $null -ne $r_homedrive -and $r_homedrive -ne "" )
        {
            $ctRDSHomeDrive++
        }

        $user =
        [PSCustomObject] @{
            samaccountname = $samaccountname
            distinguishedname = $distinguishedname
            useraccountcontrol = $useraccountcontrol
            enabled = $Enabled
            unknown = $Unknown
            lockedout = $LockedOut
            passwordnotrequired = $passwordNotRequired
            passwordneverexpires = $passwordNeverExpires
            cannotChangePassword = $cannotChangePassword
            passwordExpired = $passwordExpired
            primaryGroup = $primaryGroup
            primaryGroupID = $primaryGroupID
            lastlogonTimestamp = $lastlogontimestamp
            homedrive = $homedrive
            homedirectory = $homedirectory
            profilepath = $profilepath
            scriptpath = $scriptpath
            r_homedrive = $r_homedrive
            r_homedir = $r_homedir
            r_profpath = $r_profpath
            r_allowlogon = $r_allowlogon
        }
        ##$user

        ##
        ## many lists - single user object 
        ## we insert a reference to the user object into the list.
        ## not a copy
        ##

        $null = $listUser.Add( $user )
        If( -not $Enabled )
        {
            $null = $listUsersDisabled.Add( $user )
        }
        If( $Unknown )
        {
            $null = $listUsersUnknown.Add( $user )
        }
        If( $LockedOut )
        {
            $null = $listUsersLockedOut.Add( $user )
        }
        If( $passwordExpired )
        {
            $null = $listPasswordExpired.Add( $user )
        }
        If( $passwordNeverExpires )
        {
            $null = $listPasswordNeverExpires.Add( $user )
        }
        If( $passwordNotRequired )
        {
            $null = $listPasswordNotRequired.Add( $user )
        }
        If( $cannotChangePassword )
        {
            $null = $listCannotChangePassword.Add( $user )
        }
        If( $null -eq $lastlogontimestamp )
        {
            $null = $listNolastLogonTimestamp.Add( $user )
        }
        If( $null -ne $hasSIDHistory )
        {
            $null = $listHasSIDHistory.Add( $user )
        }
        If( $null -ne $homedrive )
        {
            $null = $listHomeDrive.Add( $user )
        }
        If( $primaryGroupID -ne 513 )
        {
            $null = $listPrimaryGroup.Add( $user )
        }
        If( $null -ne $r_homedrive -and $r_homedrive -ne "")
        {
            $null = $listRDSHomeDrive.Add( $user )
        }
    }

    Write-Verbose "$(Get-Date -Format G): `t`tGetDSUsers main processing done"
	
	#3.04 only process the foreign security principals for the root domain
	If($TrustedDomain -eq $Script:ForestRootDomain)
	{

		#FSP added in 3.03
		Write-Verbose "$(Get-Date -Format G): `t`tProcessing Orphaned Foreign Security Principals"
		
		#https://powershell.org/forums/topic/foreign-security-principals/

		# Get a list of FSPs
		$results = Get-ADObject -Filter { objectClass -eq "foreignSecurityPrincipal" } -Properties memberof | Sort-Object Name
		ForEach($result in $results)
		{
			
			$TranslatedName = $null
			try 
			{
				$TranslatedName = ([System.Security.Principal.SecurityIdentifier] $result.Name).Translate([System.Security.Principal.NTAccount])
			}
			catch 
			{
				$TranslatedName = "Orphaned"
				
				#only need the orphans
				$FSPGroups = ""
				ForEach($Group in $Result.MemberOf)
				{
					$FSPGroups += $Group
				}
				$obj = [PSCustomObject] @{
					Name           = $Result.Name
					TranslatedName = $TranslatedName
					Groups         = $FSPGroups
				}
				$listOrphanedFSPs.Add($obj) > $Null
			}
		}

		$ctOrphanedFSPs = $listOrphanedFSPs.Count
	}
    <#
	Write-Verbose "$(Get-Date -Format G): ctUsers                $ctUsers"
    Write-Verbose "$(Get-Date -Format G): ctUsersDisabled        $ctUsersDisabled"
    Write-Verbose "$(Get-Date -Format G): ctUsersUnknown         $ctUsersUnknown"
    Write-Verbose "$(Get-Date -Format G): ctUsersLockedOut       $ctUsersLockedOut"
    Write-Verbose "$(Get-Date -Format G): ctPasswordExpired      $ctPasswordExpired"
    Write-Verbose "$(Get-Date -Format G): ctPasswordNeverExpires $ctPasswordNeverExpires"
    Write-Verbose "$(Get-Date -Format G): ctPasswordNotRequired  $ctPasswordNotRequired"
    Write-Verbose "$(Get-Date -Format G): ctCannotChangePassword $ctCannotChangePassword"
    Write-Verbose "$(Get-Date -Format G): ctNolastLogonTimestamp $ctNolastLogonTimestamp"
    Write-Verbose "$(Get-Date -Format G): ctHasSIDHistory        $ctHasSIDHistory"
    Write-Verbose "$(Get-Date -Format G): ctHomeDrive            $ctHomeDrive"
    Write-Verbose "$(Get-Date -Format G): ctPrimaryGroup         $ctPrimaryGroup"
    Write-Verbose "$(Get-Date -Format G): ctRDSHomeDrive         $ctRDSHomeDrive"

    Write-Verbose "$(Get-Date -Format G): ctActiveUsers                $ctActiveUsers"
    Write-Verbose "$(Get-Date -Format G): ctActivePasswordExpired      $ctActivePasswordExpired"
    Write-Verbose "$(Get-Date -Format G): ctActivePasswordNeverExpires $ctActivePasswordNeverExpires"
    Write-Verbose "$(Get-Date -Format G): ctActivePasswordNotRequired  $ctActivePasswordNotRequired"
    Write-Verbose "$(Get-Date -Format G): ctActiveCannotChangePassword $ctActiveCannotChangePassword"
    Write-Verbose "$(Get-Date -Format G): ctActiveNolastLogonTimestamp $ctActiveNolastLogonTimestamp"

    Write-Verbose "$(Get-Date -Format G): `t`tGetDSUsers end"
    Write-Verbose "$(Get-Date -Format G): `t`tFormat numbers into strings"
	#>

    ## I pre-format the numbers because all 3 of the output formats were doing
    ## their own formatting, leading to some minor display inconsistencies. I
	## actually don't even know if it's worth it. But the variable names are
	## carefully chosen to be predictable, to make it easy to maintain the output
	## formating blocks.

    [String] $fs                            = '{0,7:N0}'  ## FormatString

    [String] $strUsers                      = $fs -f $ctUsers
    [String] $strUsersDisabled              = $fs -f $ctUsersDisabled
    [String] $strUsersUnknown               = $fs -f $ctUsersUnknown
    [String] $strUsersLockedOut             = $fs -f $ctUsersLockedOut
    [String] $strPasswordExpired            = $fs -f $ctPasswordExpired
    [String] $strPasswordNeverExpires       = $fs -f $ctPasswordNeverExpires
    [String] $strPasswordNotRequired        = $fs -f $ctPasswordNotRequired
    [String] $strCannotChangePassword       = $fs -f $ctCannotChangePassword
    [String] $strNolastLogonTimestamp       = $fs -f $ctNolastLogonTimestamp
    [String] $strHasSIDHistory              = $fs -f $ctHasSIDHistory
    [String] $strHomeDrive                  = $fs -f $ctHomeDrive
    [String] $strPrimaryGroup               = $fs -f $ctPrimaryGroup
    [String] $strRDSHomeDrive               = $fs -f $ctRDSHomeDrive
    [String] $strOrphanedFSPs               = $fs -f $ctOrphanedFSPs

    [String] $strActiveUsers                = $fs -f $ctActiveUsers
    [String] $strActivePasswordExpired      = $fs -f $ctActivePasswordExpired
    [String] $strActivePasswordNeverExpires = $fs -f $ctActivePasswordNeverExpires
    [String] $strActivePasswordNotRequired  = $fs -f $ctActivePasswordNotRequired
    [String] $strActiveCannotChangePassword = $fs -f $ctActiveCannotChangePassword
    [String] $strActiveNoLastLogonTimestamp = $fs -f $ctActiveNoLastLogonTimestamp

    [String] $fs                            = '{0,6:N2}% of total users'  ## FormatString

##  [String] $pctUsers                      = $fs -f $ctUsers
    [String] $pctUsersDisabled              = $fs -f ( ( $ctUsersDisabled        / $ctUsers ) * 100 )
    [String] $pctUsersUnknown               = $fs -f ( ( $ctUsersUnknown         / $ctUsers ) * 100 )
    [String] $pctUsersLockedOut             = $fs -f ( ( $ctUsersLockedOut       / $ctUsers ) * 100 )
    [String] $pctPasswordExpired            = $fs -f ( ( $ctPasswordExpired      / $ctUsers ) * 100 )
    [String] $pctPasswordNeverExpires       = $fs -f ( ( $ctPasswordNeverExpires / $ctUsers ) * 100 )
    [String] $pctPasswordNotRequired        = $fs -f ( ( $ctPasswordNotRequired  / $ctUsers ) * 100 )
    [String] $pctCannotChangePassword       = $fs -f ( ( $ctCannotChangePassword / $ctUsers ) * 100 )
    [String] $pctNolastLogonTimestamp       = $fs -f ( ( $ctNolastLogonTimestamp / $ctUsers ) * 100 )
    [String] $pctHasSIDHistory              = $fs -f ( ( $ctHasSIDHistory        / $ctUsers ) * 100 )
    [String] $pctHomeDrive                  = $fs -f ( ( $ctHomeDrive            / $ctUsers ) * 100 )
    [String] $pctPrimaryGroup               = $fs -f ( ( $ctPrimaryGroup         / $ctUsers ) * 100 )
    [String] $pctRDSHomeDrive               = $fs -f ( ( $ctRDSHomeDrive         / $ctUsers ) * 100 )

    [String] $pctActiveUsers                = $fs -f ( ( $ctActiveUsers          / $ctUsers ) * 100 )

    [String] $fs                            = '{0,6:N2}% of active users'  ## FormatString

    [String] $pctActivePasswordExpired      = $fs -f ( ( $ctActivePasswordExpired      / $ctActiveUsers ) * 100 )
    [String] $pctActivePasswordNeverExpires = $fs -f ( ( $ctActivePasswordNeverExpires / $ctActiveUsers ) * 100 )
    [String] $pctActivePasswordNotRequired  = $fs -f ( ( $ctActivePasswordNotRequired  / $ctActiveUsers ) * 100 )
    [String] $pctActiveCannotChangePassword = $fs -f ( ( $ctActiveCannotChangePassword / $ctActiveUsers ) * 100 )
    [String] $pctActiveNoLastLogonTimestamp = $fs -f ( ( $ctActiveNoLastLogonTimestamp / $ctActiveUsers ) * 100 )

    If( $Text )
    {
        lx 0 'All Users'
        lx 1 'Total Users                             ' $strUsers
        lx 1 'Who are unknown*                        ' $strUsersUnknown         ', ' $pctUsersUnknown
        lx 1 'Who are disabled                        ' $strUsersDisabled        ', ' $pctUsersDisabled
        lx 1 'Who are locked out                      ' $strUsersLockedOut       ', ' $pctUsersLockedOut
        lx 1 'With password expired                   ' $strPasswordExpired      ', ' $pctPasswordExpired
        lx 1 'With password never expires             ' $strPasswordNeverExpires ', ' $pctPasswordNeverExpires
        lx 1 'With password not required              ' $strPasswordNotRequired  ', ' $pctPasswordNotRequired
        lx 1 'Who cannot change password              ' $strCannotChangePassword ', ' $pctCannotChangePassword
        lx 1 'Who have never logged on                ' $strNolastLogonTimestamp ', ' $pctNolastLogonTimestamp
        lx 1 'Who have SID history                    ' $strHasSIDHistory        ', ' $pctHasSIDHistory
        lx 1 'Who have a homedrive                    ' $strHomeDrive            ', ' $pctHomeDrive
        lx 1 'Who have a primary group                ' $strPrimaryGroup         ', ' $pctPrimaryGroup
        lx 1 'Who have a RDS homedrive                ' $strRDSHomeDrive         ', ' $pctRDSHomeDrive
		If($TrustedDomain -eq $Script:ForestRootDomain)
		{
			lx 1 'Orphaned Foreign Security Principals    ' $strOrphanedFSPs         ', ' " N/A"
		}
        lx 0
        lx 1 '* Unknown users are user accounts with no UserAccountControl property.'
        lx 1 '  This should not occur.'
        If( $Script:DARights -eq $false )
        {
            lx 1 "  You should re-run the script with Domain Admin rights in $TrustedDomain." ## -options $htmlBold
        }
        Else
        {
            lx 1 "  This may be because of a permissions issue if $TrustedDomain is a Trusted Forest." ## -options $htmlBold
        }
        lx 0
        lx 1 'Active Users                ' $strActiveUsers                ', ' $pctActiveUsers
        lx 1 'With password expired       ' $strActivePasswordExpired      ', ' $pctActivePasswordExpired
        lx 1 'With password never expires ' $strActivePasswordNeverExpires ', ' $pctActivePasswordNeverExpires
        lx 1 'With password not required  ' $strActivePasswordNotRequired  ', ' $pctActivePasswordNotRequired
        lx 1 'Who cannot change password  ' $strActiveCannotChangePassword ', ' $pctActiveCannotChangePassword
        lx 1 'Who have never logged on    ' $strActiveNolastLogonTimestamp ', ' $pctActiveNolastLogonTimestamp
		Line 0 ""
    }
	If($HTML)
	{
		Write-Verbose "$(Get-Date -Format G): `t`tBuild table for All Users"
		WriteHTMLLine 3 0 'All Users'
		#V3.00 pre-allocate rowdata
		## $rowdata = @()
		$rowdata = New-Object System.Array[] 12

		$rowdata[ 0 ] = @(
			'Who are unknown*', $htmlsb,
			$strUsersUnknown,   ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctUsersUnknown,   $htmlwhite
		)

		$rowdata[ 1 ] = @(
			'Who are disabled', $htmlsb,
			$strUsersDisabled,  ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctUsersDisabled,  $htmlwhite
		)

		$rowdata[ 2 ] = @(
			'Who are locked out', $htmlsb,
			$strUsersLockedOut,   ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctUsersLockedOut,   $htmlwhite
		)

		$rowdata[ 3 ]= @(
			'With password expired', $htmlsb,
			$strPasswordExpired,     ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctPasswordExpired,     $htmlwhite
		)

		$rowdata[ 4 ] = @(
			'With password never expires', $htmlsb,
			$strPasswordNeverExpires,      ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctPasswordNeverExpires,      $htmlwhite
		)

		$rowdata[ 5 ] = @(
			'With password not required', $htmlsb,
			$strPasswordNotRequired,      ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctPasswordNotRequired,      $htmlwhite
		)

		$rowdata[ 6 ] = @(
			'Who cannot change password', $htmlsb,
			$strCannotChangePassword,     ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctCannotChangePassword,     $htmlwhite
		)

		$rowdata[ 7 ] = @(
			'Who have never logged on', $htmlsb,
			$strNolastLogonTimestamp,   ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctNolastLogonTimestamp,   $htmlwhite
		)

		$rowdata[ 8 ] = @(
			'Who have SID history', $htmlsb,
			$strHasSIDHistory,      ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctHasSIDHistory,      $htmlwhite
		)

		$rowdata[ 9 ] = @(
			'Who have a homedrive', $htmlsb,
			$strHomeDrive,          ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctHomeDrive,          $htmlwhite
		)

		$rowdata[ 10 ] = @(
			'Who have a primary group', $htmlsb,
			$strPrimaryGroup,           ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctPrimaryGroup,           $htmlwhite
		)

		$rowdata[ 11 ] = @(
			'Who have a RDS homedrive', $htmlsb,
			$strRDSHomeDrive,           ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctRDSHomeDrive,           $htmlwhite
		)

		If($TrustedDomain -eq $Script:ForestRootDomain)
		{
			$rowdata[ 11 ] = @(
				'Orphaned Foreign Security Principals', $htmlsb,
				 $strOrphanedFSPs, ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
				"N/A", $htmlwhite
			)
		}
		
		$columnWidths = @( '300px', '75px', '125px' )
		$columnHeaders = @(
			'Total Users', $htmlsb,
			$strUsers,     ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			'',            $htmlwhite
		)

		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '500'
		$rowData = $null

		WriteHTMLLine 0 0 "* Unknown users are user accounts with no UserAccountControl property." -options $htmlBold
		If($Script:DARights -eq $False)
		{
			WriteHTMLLine 0 0 "* Rerun the script with Domain Admin rights in $($ADForest)." -options $htmlBold
		}
		Else
		{
			WriteHTMLLine 0 0 "* This may be a permissions issue if this is a Trusted Forest." -options $htmlBold
		}
		
		Write-Verbose "$(Get-Date -Format G): `t`tBuild table for Active Users"
		WriteHTMLLine 3 0 "Active Users"

		#V3.00 pre-allocate rowdata
		## $rowdata = @()
		$rowdata = New-Object System.Array[] 6

		$rowdata[ 0 ] = @(
			'Total Active Users', $htmlsb,
			$strActiveUsers,      ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctActiveUsers,      $htmlwhite
		)

		$rowdata[ 1 ] = @(
			'With password expired',   $htmlsb,
			$strActivePasswordExpired, ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctActivePasswordExpired, $htmlwhite
		)

		$rowdata[ 2 ] = @(
			'With password never expires',  $htmlsb,
			$strActivePasswordNeverExpires, ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctActivePasswordNeverExpires, $htmlwhite
		)

		$rowdata[ 3 ] = @(
			'With password not required',  $htmlsb,
			$strActivePasswordNotRequired, ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctActivePasswordNotRequired, $htmlwhite
		)

		$rowdata[ 4 ] = @(
			'Who cannot change password',   $htmlsb,
			$strActiveCannotChangePassword, ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctActiveCannotChangePassword, $htmlwhite
		)

		$rowdata[ 5 ] = @(
			'Who have never logged on',     $htmlsb,
			$strActiveNolastLogonTimestamp, ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			$pctActiveNolastLogonTimestamp, $htmlwhite
		)

		$columnWidths = @( '300px', '75px', '150px' )
		$columnHeaders = @(
			'Total Users', $htmlsb,
			$strUsers,     ( $htmlwhite -bor $global:htmlAlignRight ), #3.06 add alignright
			'',            $htmlwhite
		)

		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '500'
		WriteHTMLLine 0 0 ''
	}
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): `t`tBuild table for All Users"
		WriteWordLine 3 0 'All Users'
		$TableRange   = $Script:doc.Application.Selection.Range
		[int]$Columns = 3
		If($TrustedDomain -eq $Script:ForestRootDomain)
		{
			[int]$Rows = 14
		}
		Else
		{
			[int]$Rows = 13
		}
		$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.Style = $Script:MyHash.Word_TableGrid

		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = 'Total Users'
		$Table.Cell(1,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(1,2).Range.Text = $strUsers

		$Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(2,1).Range.Font.Bold = $True
		$Table.Cell(2,1).Range.Text = 'Disabled users'
		$Table.Cell(2,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(2,2).Range.Text = $strUsersDisabled
		$Table.Cell(2,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(2,3).Range.Text = $pctUsersDisabled

		$Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(3,1).Range.Font.Bold = $True
		$Table.Cell(3,1).Range.Text = "Unknown users*"
		$Table.Cell(3,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(3,2).Range.Text = $strUsersUnknown
		$Table.Cell(3,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(3,3).Range.Text = $pctUsersUnknown

		$Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(4,1).Range.Font.Bold = $True
		$Table.Cell(4,1).Range.Text = "Locked out users"
		$Table.Cell(4,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(4,2).Range.Text = $strUsersLockedOut
		$Table.Cell(4,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(4,3).Range.Text = $pctUsersLockedOut

		$Table.Cell(5,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(5,1).Range.Font.Bold = $True
		$Table.Cell(5,1).Range.Text = "Password expired"
		$Table.Cell(5,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(5,2).Range.Text = $strPasswordExpired
		$Table.Cell(5,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(5,3).Range.Text = $pctPasswordExpired

		$Table.Cell(6,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(6,1).Range.Font.Bold = $True
		$Table.Cell(6,1).Range.Text = "Password never expires"
		$Table.Cell(6,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(6,2).Range.Text = $strPasswordNeverExpires
		$Table.Cell(6,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(6,3).Range.Text = $pctPasswordNeverExpires

		$Table.Cell(7,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(7,1).Range.Font.Bold = $True
		$Table.Cell(7,1).Range.Text = "Password not required"
		$Table.Cell(7,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(7,2).Range.Text = $strPasswordNotRequired
		$Table.Cell(7,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(7,3).Range.Text = $pctPasswordNotRequired

		$Table.Cell(8,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(8,1).Range.Font.Bold = $True
		$Table.Cell(8,1).Range.Text = "Can't change password"
		$Table.Cell(8,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(8,2).Range.Text = $strCannotChangePassword
		$Table.Cell(8,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(8,3).Range.Text = $pctCannotChangePassword

		$Table.Cell(9,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(9,1).Range.Font.Bold = $True
		$Table.Cell(9,1).Range.Text = "Who have not logged on"
		$Table.Cell(9,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(9,2).Range.Text = $strNolastLogonTimestamp
		$Table.Cell(9,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(9,3).Range.Text = $pctNolastLogonTimestamp

		$Table.Cell(10,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(10,1).Range.Font.Bold = $True
		$Table.Cell(10,1).Range.Text = "With SID History"
		$Table.Cell(10,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(10,2).Range.Text = $strHasSIDHistory
		$Table.Cell(10,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(10,3).Range.Text = $pctHasSIDHistory

		$Table.Cell(11,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(11,1).Range.Font.Bold = $True
		$Table.Cell(11,1).Range.Text = "HomeDrive users"
		$Table.Cell(11,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(11,2).Range.Text = $strHomeDrive
		$Table.Cell(11,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(11,3).Range.Text = $pctHomeDrive

		$Table.Cell(12,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(12,1).Range.Font.Bold = $True
		$Table.Cell(12,1).Range.Text = "PrimaryGroup users"
		$Table.Cell(12,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(12,2).Range.Text = $strPrimaryGroup
		$Table.Cell(12,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(12,3).Range.Text = $pctPrimaryGroup

		$Table.Cell(13,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(13,1).Range.Font.Bold = $True
		$Table.Cell(13,1).Range.Text = "RDS HomeDrive users"
		$Table.Cell(13,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(13,2).Range.Text = $strRDSHomeDrive
		$Table.Cell(13,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(13,3).Range.Text = $pctRDSHomeDrive

		If($TrustedDomain -eq $Script:ForestRootDomain)
		{
			$Table.Cell(14,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(14,1).Range.Font.Bold = $True
			$Table.Cell(14,1).Range.Text = "Orphaned Foreign Security Principals"
			$Table.Cell(14,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
			$Table.Cell(14,2).Range.Text = $strOrphanedFSPs
			$Table.Cell(14,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
			$Table.Cell(14,3).Range.Text = "N/A"
		}

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitContent)

		#Return focus back to document
		$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null

		WriteWordLine 0 0 "* Unknown users are user accounts with no Enabled property." "" $Null 8 $False $True
		If($Script:DARights -eq $False)
		{
			WriteWordLine 0 0 "* Rerun the script with Domain Admin rights in $($ADForest)." "" $Null 8 $False $True
		}
		Else
		{
			WriteWordLine 0 0 "* This may be a permissions issue if this is a Trusted Forest." "" $Null 8 $False $True
		}
		
		Write-Verbose "$(Get-Date -Format G): `t`tBuild table for Active Users"
		WriteWordLine 3 0 "Active Users"
		$TableRange   = $Script:doc.Application.Selection.Range
		[int]$Columns = 3
		[int]$Rows = 6
		$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.Style = $Script:MyHash.Word_TableGrid

		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Total Active Users"
		$Table.Cell(1,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(1,2).Range.Text = $strActiveUsers
		$Table.Cell(1,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(2,3).Range.Text = $pctActiveUsers

		$Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(2,1).Range.Font.Bold = $True
		$Table.Cell(2,1).Range.Text = "Password expired"
		$Table.Cell(2,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(2,2).Range.Text = $strActivePasswordExpired
		$Table.Cell(2,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(2,3).Range.Text = $pctActivePasswordExpired

		$Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(3,1).Range.Font.Bold = $True
		$Table.Cell(3,1).Range.Text = "Password never expires"
		$Table.Cell(3,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(3,2).Range.Text = $strActivePasswordNeverExpires
		$Table.Cell(3,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(3,3).Range.Text = $pctActivePasswordNeverExpires

		$Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(4,1).Range.Font.Bold = $True
		$Table.Cell(4,1).Range.Text = "Password not required"
		$Table.Cell(4,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(4,2).Range.Text = $strActivePasswordNotRequired
		$Table.Cell(4,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(4,3).Range.Text = $pctActivePasswordNotRequired

		$Table.Cell(5,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(5,1).Range.Font.Bold = $True
		$Table.Cell(5,1).Range.Text = "Can't change password"
		$Table.Cell(5,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(5,2).Range.Text = $strActiveCannotChangePassword
		$Table.Cell(5,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(5,3).Range.Text = $pctActiveCannotChangePassword

		$Table.Cell(6,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(6,1).Range.Font.Bold = $True
		$Table.Cell(6,1).Range.Text = "No lastLogonTimestamp"
		$Table.Cell(6,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(6,2).Range.Text = $strActiveNolastLogonTimestamp
		$Table.Cell(6,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(6,3).Range.Text = $pctActiveNolastLogonTimestamp

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitContent)

		#Return focus back to document
		$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null

		#put computer info on a separate page
		$Script:selection.InsertNewPage()
	}

	If($IncludeUserInfo -eq $True)
	{
		If($ctUsersDisabled -gt 0)
		{
			OutputUserInfo $listUsersDisabled 'Disabled users'
		}

		If($ctUsersUnknown -gt 0)
		{
			OutputUserInfo $listUsersUnknown 'Unknown users'
		}

		If($ctUsersLockedOut -gt 0)
		{
			OutputUserInfo $listUsersLockedOut 'Locked out users'
		}

		If($ctPasswordExpired -gt 0)
		{
			OutputUserInfo $listPasswordExpired 'All users with password expired'
		}

		If($ctPasswordNeverExpires -gt 0)
		{
			OutputUserInfo $listPasswordNeverExpires 'All users whose password never expires'
		}

		If($ctPasswordNotRequired -gt 0)
		{
			OutputUserInfo $listPasswordNotRequired 'All users with password not required'
		}

		If($ctCannotChangePassword -gt 0)
		{
			OutputUserInfo $listCannotChangePassword 'All users who cannot change password'
		}

		If($ctHasSIDHistory -gt 0)
		{
			OutputUserInfo $listHasSIDHistory 'All users with SID History'
		}

		If($ctHomeDrive -gt 0)
		{
			OutputHDUserInfo $listHomeDrive 'All users with HomeDrive set in ADUC'
		}
	
		If($ctPrimaryGroup -gt 0)
		{
			OutputPGUserInfo $listPrimaryGroup 'All users whose Primary Group is not Domain Users'
		}

		If($ctRDSHomeDrive -gt 0)
		{
			OutputRDSHDUserInfo $listRDSHomeDrive 'All users with RDS HomeDrive set in ADUC'
		}

		If($ctOrphanedFSPs -gt 0)
		{
			OutputFSPUserInfo $listOrphanedFSPs 'Orphaned Foreign Security Principals'
		}

	}

	$script:MaxPasswordAge = $null

    #$script_end = Get-Date
    #$script_delta = $script_end - $script_begin
	#$elapsed = 'Elapsed: ' + $script_delta.Hours.ToString() + '.' + $script_delta.Minutes.ToString() + '.' + $script_delta.Seconds.ToString()
	#Write-Verbose "$(Get-Date -Format G):`tEnd GetDSusers TrustedDomain $TrustedDomain, Elapsed $elapsed"
}

Function ProcessMiscDataByDomain
{
	Write-Verbose "$(Get-Date -Format G): Writing miscellaneous data by domain"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Miscellaneous Data by Domain"
	}
	If($Text)
	{
		Line 0 "///  Miscellaneous Data by Domain  \\\"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Miscellaneous Data by Domain&nbsp;&nbsp;\\\"
	}
	
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing misc data for domain $($Domain)"

		$DomainInfo = Get-ADDomain -Identity $Domain -EA 0 

		If( !$? )
		{
			$txt = "Error retrieving domain data for domain $Domain"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt '' $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If( $null -eq $DomainInfo )
		{
			$txt = "No Domain data was retrieved for domain $Domain"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			If($Text)
			{
				Line 0 $txt
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}

			Continue
		}

		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}

		$txt = $Domain
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt += ' (Forest Root)'
		}

		If($MSWORD -or $PDF)
		{
			WriteWordLine 2 0 $txt
		}
		If($Text)
		{
			Line 0 "///  $($txt)  \\\"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
		}

		Write-Verbose "$(Get-Date -Format G): `t`tGathering user misc data #2"
		
		getDSUsers $Domain

		Get-ComputerCountByOS $Domain

		$First = $False
	}

	$Script:Domains = $Null
} ## end ProcessMiscDataByDomain

Function OutputUserInfo
{
	Param
	(
		[Object[]] $Users, 
		[String] $title
	)
	
	Write-Verbose "$(Get-Date -Format G): `t`t`t`tOutput $($title)"
	$Users = $Users | Sort-Object samAccountName
	
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $UsersWordTable = @()

		WriteWordLine 4 0 $title

		ForEach($User in $Users)
		{
			$WordTableRowHash = @{ 
				SamAccountName = $User.SamAccountName; 
				DN = $User.DistinguishedName
			}

			## Add the hash to the array
			$UsersWordTable += $WordTableRowHash
		}
		
		## Add the table to the document, using the hashtable
		If($UsersWordTable.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $UsersWordTable `
			-Columns SamAccountName, DN `
			-Headers "SamAccountName", "DistinguishedName" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 350;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($Text)
	{
		Line 0 $title
		Line 0 ""
		Line 1 "samAccountName            DistinguishedName"
		Line 1 "=============================================================================================================================================="
		###### "1234512345123451234512345

		ForEach($User in $Users)
		{
			Line 1 ( "{0,-25} {1,-116}" -f $User.samAccountName, $User.DistinguishedName )
		}
		Line 0 ""
	}
	If($HTML)
	{
		#V3.00 pre-allocate rowdata
		## $rowdata = @()
		If( $Users -and $Users.Length -gt 0 )
		{
			$arrayLength = $Users.Length
		}
		Else
		{
			$arrayLength = 0
		}
		## wv "***** OutputUserInfo: users.length $( $arrayLength ), gettype $( $Users.GetType().FullName ), title = '$title'"

		WriteHTMLLine 4 0 ( $title + ' (' + $arrayLength.ToString() + ')' )
		$rowdata  = New-Object System.Array[] $arrayLength
		$rowIndex = 0

		ForEach($User in $Users)
		{
			$rowdata[ $rowIndex ] = @(
				$User.SamAccountName,    $htmlwhite,
				$User.DistinguishedName, $htmlwhite
			)
			$rowIndex++
		}
		
		$columnWidths  = @( '150px', '350px' )
		$columnHeaders = @(
			'SamAccountName',    $htmlsb,
			'DistinguishedName', $htmlsb
		)

		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '500'
		WriteHTMLLine 0 0 ''
	}
} ## end Function OutputUserInfo

Function OutputHDUserInfo
{
	#new for 2.16
	Param
	(
		[Object[]] $Users,
		[string] $title
	)
	
	Write-Verbose "$(Get-Date -Format G): `t`t`t`tOutput $($title)"
	$Users = $Users | Sort-Object samAccountName
	
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $UsersWordTable = @()

		WriteWordLine 4 0 $title

		ForEach($User in $Users)
		{
			$WordTableRowHash = @{ 
				SamAccountName = $User.SamAccountName; 
				DN = $User.DistinguishedName;
				HomeDrive = $User.HomeDrive;
				HomeDir = $User.HomeDirectory;
				ProfilePath = $User.ProfilePath;
				ScriptPath = $User.ScriptPath
			}

			## Add the hash to the array
			$UsersWordTable += $WordTableRowHash
		}
		
		## Add the table to the document, using the hashtable
		If($UsersWordTable.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $UsersWordTable `
			-Columns SamAccountName, DN, HomeDrive, HomeDir, ProfilePath, ScriptPath `
			-Headers "SamAccountName", "DistinguishedName", "Home drive", "Home folder", "Profile path", "Login script" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 100;
			$Table.Columns.Item(2).Width = 110;
			$Table.Columns.Item(3).Width = 35;
			$Table.Columns.Item(4).Width = 85;
			$Table.Columns.Item(5).Width = 85;
			$Table.Columns.Item(6).Width = 85;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ''
		}
	}
	If($Text)
	{
		Line 0 $title
		Line 0 ''

		ForEach($User in $Users)
		{
			Line 1 "SamAccountName`t`t: " $User.samAccountName
			Line 1 "DistinguishedName`t: " $User.DistinguishedName
			Line 1 "Home drive`t`t: " $User.HomeDrive
			Line 1 "Home folder`t`t: " $User.HomeDirectory
			Line 1 "Profile path`t`t: " $User.ProfilePath
			Line 1 "Login script`t`t: " $User.ScriptPath
			Line 0 ""
		}
	}
	If($HTML)
	{
		#V3.00 pre-allocate rowdata
		## $rowdata = @()
		If( $Users -and $Users.Length -gt 0 )
		{
			$arrayLength = $Users.Length
		}
		Else
		{
			$arrayLength = 0
		}

		WriteHTMLLine 4 0 ( $title + ' (' + $arrayLength.ToString() + ')' )
		$rowdata  = New-Object System.Array[] $arrayLength
		$rowIndex = 0
		
		ForEach($User in $Users)
		{
			$rowdata[ $rowIndex ] = @(
				$User.SamAccountName,    $htmlwhite,
				$User.DistinguishedName, $htmlwhite,
				$User.HomeDrive,         $htmlwhite,
				$User.HomeDirectory,     $htmlwhite,
				$User.ProfilePath,       $htmlwhite,
				$User.ScriptPath,        $htmlwhite
			)
			$rowIndex++
		}

		$columnWidths  = @( '100px', '100px', '75px', '75px', '75px', '75px' )
		$columnHeaders = @(
			'SamAccountName',    $htmlsb,
			'DistinguishedName', $htmlsb,
			'Home Drive',        $htmlsb,
			'Home folder',       $htmlsb,
			'Profile path',      $htmlsb,
			'Login script',      $htmlsb
		)

		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '500'
		WriteHTMLLine 0 0

		$rowdata = $null
	}
} ## end Function OutputHDUserInfo

Function OutputPGUserInfo
{
	#new for 2.16
	Param
	(
		[Object[]] $Users,
		[string] $title
	)
	
	Write-Verbose "$(Get-Date -Format G): `t`t`t`tOutput $($title)"
	$Users = $Users | Sort-Object samAccountName
	
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $UsersWordTable = @()

		WriteWordLine 4 0 $title

		ForEach($User in $Users)
		{
			$WordTableRowHash = @{
				SamAccountName = $User.SamAccountName; 
				DN = $User.DistinguishedName;
				PG = $User.PrimaryGroup
			}

			## Add the hash to the array
			$UsersWordTable += $WordTableRowHash
		}
		
		## Add the table to the document, using the hashtable
		If($UsersWordTable.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $UsersWordTable `
			-Columns SamAccountName, DN, PG `
			-Headers "SamAccountName", "DistinguishedName", "Primary Group" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 100;
			$Table.Columns.Item(2).Width = 200;
			$Table.Columns.Item(3).Width = 200;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($Text)
	{
		Line 0 $title
		Line 0 ""

		ForEach($User in $Users)
		{
			Line 1 "samAccountName`t`t: " $User.samAccountName
			Line 1 "DistinguishedName`t: " $User.DistinguishedName
			Line 1 "Primary Group`t`t: " $User.PrimaryGroup
			Line 0 ""
		}
	}
	If($HTML)
	{
		#V3.00 pre-allocate rowdata
		## $rowdata = @()
		If( $Users -and $Users.Length -gt 0 )
		{
			$arrayLength = $Users.Length
		}
		Else
		{
			$arrayLength = 0
		}

		WriteHTMLLine 4 0 ( $title + ' (' + $arrayLength.ToString() + ')' )
		$rowdata  = New-Object System.Array[] $arrayLength
		$rowIndex = 0
		
		ForEach($User in $Users)
		{
			$rowdata[ $rowIndex ] = @(
				$User.SamAccountName,    $htmlwhite,
				$User.DistinguishedName, $htmlwhite,
				$User.PrimaryGroup,      $htmlwhite
			)
			$rowIndex++
		}

		$columnWidths  = @( '100px', '200px', '200px' )
		$columnHeaders = @(
			'SamAccountName',    $htmlsb,
			'DistinguishedName', $htmlsb,
			'Primary Group',     $htmlsb
		)

		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '500'
		WriteHTMLLine 0 0 ''

		$rowdata = $null
	}
} ## end Function OutputPGUserInfo

Function OutputRDSHDUserInfo
{
	#new for 2.16
	Param
	(
		[Object[]] $Users, 
		[string] $title
	)

	Write-Verbose "$(Get-Date -Format G): `t`t`t`tOutput $($title)"
	$Users = $Users | Sort-Object samAccountName
	
	If( $Users -and $Users.Length -gt 0 )
	{
		$arrayLength = $Users.Length
	}
	Else
	{
		$arrayLength = 0
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $UsersWordTable = @()

		WriteWordLine 4 0 ( $title + ' (' + $arrayLength.ToString() + ')' )

		ForEach($User in $Users)
		{
			$WordTableRowHash = @{
				SamAccountName = $User.SamAccountName; 
				DN = $User.DistinguishedName;
				HomeDrive = $User.r_homedrive;
				HomeDir = $User.r_homedir;
				ProfilePath = $User.r_profpath;
				AllowLogon = $User.r_allowlogon
			}

			## Add the hash to the array
			$UsersWordTable += $WordTableRowHash
		}
		
		## Add the table to the document, using the hashtable
		If($UsersWordTable.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $UsersWordTable `
			-Columns SamAccountName, DN, HomeDrive, HomeDir, ProfilePath, ALlowLogon `
			-Headers "SamAccountName", "DistinguishedName", "RDS Home drive", "RDS Home folder", "RDS Profile path", "Allow Logon" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 100;
			$Table.Columns.Item(2).Width = 140;
			$Table.Columns.Item(3).Width = 35;
			$Table.Columns.Item(4).Width = 75;
			$Table.Columns.Item(5).Width = 90;
			$Table.Columns.Item(6).Width = 60;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($Text)
	{
		Line 0 ( $title + ' (' + $arrayLength.ToString() + ')' )
		Line 0 ""

		ForEach($User in $Users)
		{
			Line 1 "SamAccountName`t`t: " $User.samAccountName
			Line 1 "DistinguishedName`t: " $User.DistinguishedName
			Line 1 "RDS Home drive`t`t: " $User.r_homedrive
			Line 1 "RDS Home folder`t`t: " $User.r_homedir
			Line 1 "RDS Profile path`t: " $User.r_profpath
			Line 1 "Allow Logon`t`t: " $User.r_allowlogon
			Line 0 ""
		}
	}
	If($HTML)
	{
		#V3.00 pre-allocate rowdata
		## $rowdata = @()
		WriteHTMLLine 4 0 ( $title + ' (' + $arrayLength.ToString() + ')' )
		$rowdata  = New-Object System.Array[] $arrayLength
		$rowIndex = 0
		
		ForEach($User in $Users)
		{
			$rowdata[ $rowIndex ] = @(
				$User.SamAccountName,$htmlwhite,
				$User.DistinguishedName,$htmlwhite,
				$User.r_homedrive,$htmlwhite,
				$User.r_homedir,$htmlwhite,
				$User.r_profpath,$htmlwhite,
				$User.r_allowlogon,$htmlwhite
			)
			$rowIndex++
		}

		$columnWidths  = @( '100px', '100px', '75px', '75px', '90px', '60px' )
		$columnHeaders = @(
			'SamAccountName',    $htmlsb,
			'DistinguishedName', $htmlsb,
			'RDS Home Drive',    $htmlsb,
			'RDS Home folder',   $htmlsb,
			'RDS Profile path',  $htmlsb,
			'Allow Logon',       $htmlsb
		)

		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '500'
		WriteHTMLLine 0 0 ''

		$rowdata = $null
	}
} ## end Function OutputRDSHDUserInfo

Function OutputFSPUserInfo
{
	#new for 3.03
	Param
	(
		[Object[]] $Users, 
		[string] $title
	)

	Write-Verbose "$(Get-Date -Format G): `t`t`t`tOutput $($title)"
	$Users = $Users | Sort-Object Name
	
	If( $Users -and $Users.Length -gt 0 )
	{
		$arrayLength = $Users.Length
	}
	Else
	{
		$arrayLength = 0
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $UsersWordTable = @()

		WriteWordLine 4 0 ( $title + ' (' + $arrayLength.ToString() + ')' )

		ForEach($User in $Users)
		{
			$WordTableRowHash = @{
				Name   = $User.Name; 
				Groups = $User.Groups
			}

			## Add the hash to the array
			$UsersWordTable += $WordTableRowHash
		}
		
		## Add the table to the document, using the hashtable
		If($UsersWordTable.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $UsersWordTable `
			-Columns Name, Groups `
			-Headers "Name", "Groups" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 250;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($Text)
	{
		Line 0 ( $title + ' (' + $arrayLength.ToString() + ')' )
		Line 0 ""

		ForEach($User in $Users)
		{
			Line 1 "Name`t: " $User.Name
			Line 1 "Groups`t: " $User.Groups
			Line 0 ""
		}
	}
	If($HTML)
	{
		#V3.00 pre-allocate rowdata
		## $rowdata = @()
		WriteHTMLLine 4 0 ( $title + ' (' + $arrayLength.ToString() + ')' )
		$rowdata  = New-Object System.Array[] $arrayLength
		$rowIndex = 0
		
		ForEach($User in $Users)
		{
			$rowdata[ $rowIndex ] = @(
				$User.Name,$htmlwhite,
				$User.Groups,$htmlwhite
			)
			$rowIndex++
		}

		$columnWidths  = @( '400', '300')
		$columnHeaders = @(
			'Name',   $htmlsb,
			'Groups', $htmlsb
		)

		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '700'
		WriteHTMLLine 0 0 ''

		$rowdata = $null
	}
} ## end Function OutputFPSUserInfo
#endregion

#region DCDNSInfo
Function ProcessDCDNSInfo
{
	$ContinueOn = $True
	If($MSWord -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 'Domain Controller DNS IP Configuration'
		If(! ($Script:DARights -and $Script:Elevated) )
		{
			WriteWordLine 0 0 "To obtain Domain Controller DNS IP Configuration, you must run this script elevated and from an account with Domain Admin rights."
			$ContinueOn = $False
		}
	}
	If($Text)
	{
		Line 0 '///  Domain Controller DNS IP Configuration  \\\'
		Line 0 ''
		If(! ($Script:DARights -and $Script:Elevated) )
		{
			Line 0 "To obtain Domain Controller DNS IP Configuration, you must run this script elevated and from an account with Domain Admin rights."
			$ContinueOn = $False
		}
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Controller DNS IP Configuration&nbsp;&nbsp;\\\"
		If(! ($Script:DARights -and $Script:Elevated) )
		{
			WriteHTMLLine 0 0 "To obtain Domain Controller DNS IP Configuration, you must run this script elevated and from an account with Domain Admin rights."
			$ContinueOn = $False
		}
	}
	
	If($ContinueOn -eq $False)
	{
		Return
	}
	## Domain Controller DNS IP Configuration
	Write-Verbose "$(Get-Date -Format G): Create Domain Controller DNS IP Configuration"
	Write-Verbose "$(Get-Date -Format G): `tAdd Domain Controller DNS IP Configuration table to doc"
	
	## sort by site then by DC
	$xDCDNSIPInfo = @( $Script:DCDNSIPInfo | Sort-Object DCSite, DCName )

	If($MSWord -or $PDF)
	{
		$ItemsWordTable = New-Object System.Collections.ArrayList

		ForEach( $Item in $xDCDNSIPInfo )
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ 
				DCName       = $Item.DCName;
				DCSite       = $Item.DCSite;
				DCIpAddress1 = $Item.DCIpAddress1;
				DCIpAddress2 = $Item.DCIpAddress2;
				DCDNS1       = $Item.DCDNS1; 
				DCDNS2       = $Item.DCDNS2; 
				DCDNS3       = $Item.DCDNS3; 
				DCDNS4       = $Item.DCDNS4
			}

			## Add the hash to the array
			$ItemsWordTable.Add($WordTableRowHash) > $Null
		}

		## Add the table to the document, using the hashtable
		If($ItemsWordTable.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns DCName, DCSite, DCIpAddress1, DCIpAddress2, DCDNS1, DCDNS2, DCDNS3, DCDNS4 `
			-Headers "DC Name", "Site", "IP Address 1", "IP Address 2", "DNS 1", "DNS 2", "DNS 3", "DNS 4" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed

			SetWordCellFormat -Collection $Table -Size 8 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 100;
			$Table.Columns.Item(2).Width = 60;
			$Table.Columns.Item(3).Width = 70;
			$Table.Columns.Item(4).Width = 70;
			$Table.Columns.Item(5).Width = 50;
			$Table.Columns.Item(6).Width = 50;
			$Table.Columns.Item(7).Width = 50;
			$Table.Columns.Item(8).Width = 50;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
	}
	If( $Text )
	{
		ForEach( $Item in $xDCDNSIPInfo )
		{
			Line 1 "DC Name`t`t: "		$Item.DCName
			Line 1 "Site Name`t: "		$Item.DCSite
			Line 1 "IP Address1`t: "	$Item.DCIpAddress1
			Line 1 "IP Address2`t: "	$Item.DCIpAddress2
			Line 1 "DNS 1`t`t: " 		$Item.DCDNS1
			Line 1 "DNS 2`t`t: "		$Item.DCDNS2
			Line 1 "DNS 3`t`t: "		$Item.DCDNS3
			Line 1 "DNS 4`t`t: "		$Item.DCDNS4
			Line 0 ''
		}
	}
	If( $HTML )
	{
		#V3.00 pre-allocate rowdata

		$columnHeaders = @(
			'DC Name',       $htmlsb,
			'Site',          $htmlsb,
			'IP Address 1',  $htmlsb,
			'IP Address 2',  $htmlsb,
			'DNS 1',         $htmlsb,
			'DNS 2',         $htmlsb,
			'DNS 3',         $htmlsb,
			'DNS 4',         $htmlsb
		)

		$XXXrowdata = New-Object System.Array[] $xDCDNSIPInfo.Length
		$XXXrowIndx = 0

		ForEach( $Item in $xDCDNSIPInfo )
		{
			$r = @(
				$Item.DCName,       $htmlwhite,
				$Item.DCSite,       $htmlwhite,
				$Item.DCIpAddress1, $htmlwhite,
				$Item.DCIpAddress2, $htmlwhite,
				$Item.DCDNS1,       $htmlwhite,
				$Item.DCDNS2,       $htmlwhite,
				$Item.DCDNS3,       $htmlwhite,
				$Item.DCDNS4,       $htmlwhite
			)
			$XXXrowdata[ $XXXrowIndx ] = $r
			$XXXrowIndx++
		}
<#
		If( $ExtraSpecialVerbose )
		{
			wv "***** ProcessDCDNSInfo: rowdata length $( $XXXrowdata.Length )"
			for( $ii = 0; $ii -lt $XXXrowdata.Length; $ii++ )
			{
				$row = $XXXrowdata[ $ii ]
				wv "***** ProcessDCDNSInfo: rowdata index $ii, type $( $row.GetType().FullName ), length $( $row.Length )"
				for( $yyy = 0; $yyy -lt $row.Length; $yyy++ )
				{
					wv "***** ProcessDCDNSInfo: row[ $yyy ] = $( $row[ $yyy ] )"
				}
				wv "***** ProcessDCDNSInfo: done"
			}
		}
#>

		FormatHTMLTable -rowArray $XXXrowdata -columnArray $columnHeaders 
		WriteHTMLLine 0 0 ''

		$XXXrowdata    = $null
		$columnHeaders = $null
	}

	Write-Verbose "$(Get-Date -Format G): Finished Create Domain Controller DNS IP Configuration"
	Write-Verbose "$(Get-Date -Format G): "
} ## end Function ProcessDCDNSInfo
#endregion

#region TimeServerInfo
Function ProcessTimeServerInfo
{
	#Domain Controller Time Server Configuration
	$ContinueOn = $True
	If($MSWord -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Domain Controller Time Server Configuration"
		If(! ($Script:DARights -and $Script:Elevated) )
		{
			WriteWordLine 0 0 "To obtain time server information, you must run this script elevated and from an account with Domain Admin rights."
			$ContinueOn = $False
		}
	}
	If($Text)
	{
		Line 0 "///  Domain Controller Time Server Configuration  \\\"
		Line 0 ""
		If(! ($Script:DARights -and $Script:Elevated) )
		{
			Line 0 "To obtain time server information, you must run this script elevated and from an account with Domain Admin rights."
			$ContinueOn = $False
		}
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Controller Time Server Configuration&nbsp;&nbsp;\\\"
		If(! ($Script:DARights -and $Script:Elevated) )
		{
			WriteHTMLLine 0 0 "To obtain time server information, you must run this script elevated and from an account with Domain Admin rights."
			$ContinueOn = $False
		}
	}
	
	If($ContinueOn -eq $False)
	{
		Return
	}
	
	Write-Verbose "$(Get-Date -Format G): Create Domain Controller Time Server Configuration"
	Write-Verbose "$(Get-Date -Format G): `tAdd Domain Controller Time Server Configuration table to doc"
	
	#sort by DC
	$xTimeServerInfo = $Script:TimeServerInfo | Sort-Object DCName
	
	If($MSWord -or $PDF )
	{
		$ItemsWordTable = New-Object System.Collections.ArrayList

		ForEach($Item in $xTimeServerInfo)
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ 
				DCName                  = $Item.DCName;
				DCTimeSource            = $Item.TimeSource;
				DCAnnounceFlags         = $Item.AnnounceFlags;
				DCMaxNegPhaseCorrection = $Item.MaxNegPhaseCorrection;
				DCMaxPosPhaseCorrection = $Item.MaxPosPhaseCorrection;
				DCNtpServer             = $Item.NtpServer;
				DCNtpType               = $Item.NtpType;
				DCSpecialPollInterval   = $Item.SpecialPollInterval;
				DCVMICTimeProvider      = $Item.VMICTimeProvider
			}

			## Add the hash to the array
			$ItemsWordTable.Add($WordTableRowHash) > $Null
		}

		## Add the table to the document, using the hashtable
		If($ItemsWordTable.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns DCName, DCTimeSource, DCAnnounceFlags, DCMaxNegPhaseCorrection, DCMaxPosPhaseCorrection, DCNtpServer, DCNtpType, DCSpecialPollInterval, DCVMICTimeProvider `
			-Headers "DC Name", "Time Source", "Announce Flags", "Max Neg Phase Correction", "Max Pos Phase Correction", "NTP Server", "Type", "Special Poll Interval", "VMIC Time Provider" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table -Size 8 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 105;
			$Table.Columns.Item(2).Width = 73;
			$Table.Columns.Item(3).Width = 45;
			$Table.Columns.Item(4).Width = 47;
			$Table.Columns.Item(5).Width = 47;
			$Table.Columns.Item(6).Width = 73;
			$Table.Columns.Item(7).Width = 33;
			$Table.Columns.Item(8).Width = 37;
			$Table.Columns.Item(9).Width = 40;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
	}
	If( $Text )
	{

		ForEach( $Item in $xTimeServerInfo )
		{
			Line 1 "DC Name`t`t`t: "            $Item.DCName
			Line 2 "Time source`t`t: "          $Item.TimeSource
			Line 2 "Announce flags`t`t: "       $Item.AnnounceFlags
			Line 2 "Max Neg Phase Correction: " $Item.MaxNegPhaseCorrection
			Line 2 "Max Pos Phase Correction: " $Item.MaxPosPhaseCorrection
			Line 2 "NTP Server`t`t: "           $Item.NtpServer
			Line 2 "Type`t`t`t: "               $Item.NtpType
			Line 2 "Special Poll Interval`t: "  $Item.SpecialPollInterval
			Line 2 "VMIC Time Provider`t: "     $Item.VMICTimeProvider
			Line 0 ''
		}
	}
	If( $HTML )
	{
		#V3.00 pre-allocate rowdata
		#$rowdata = @()
		$rowCt = 1
		If( $xTimeServerInfo -is [Array] )
		{
			$rowCt = $xTimeServerInfo.Count
		}
#wv "Domain Controller Time Server Configuration" #MBS
		$rowData = New-Object System.Array[] $rowCt
#wv "rowCt $rowCt" #MBS
		$rowIndx = 0

		ForEach( $Item in $xTimeServerInfo )
		{
#wv "rowIndx $rowIndx $( $Item.DCName )" #MBS
			$rowdata[ $rowIndx ] = @(
				$Item.DCName,                $htmlwhite,
				$Item.TimeSource,            $htmlwhite,
				$Item.AnnounceFlags,         $htmlwhite,
				$Item.MaxNegPhaseCorrection, $htmlwhite,
				$Item.MaxPosPhaseCorrection, $htmlwhite,
				$Item.NtpServer,             $htmlwhite,
				$Item.NtpType,               $htmlwhite,
				$Item.SpecialPollInterval,   $htmlwhite,
				$Item.VMICTimeProvider,      $htmlwhite
			)
			$rowIndx++
		}

		$columnWidths  = @( '100px', '70px', '45px', '45px', '45px', '75px', '40px', '40px', '40px' )
		$columnHeaders = @(
			'DC Name',                  $htmlsb,
			'Time Source',              $htmlsb,
			'Announce Flags',           $htmlsb,
			'Max Neg Phase Correction', $htmlsb,
			'Max Pos Phase Correction', $htmlsb,
			'NTP Server',               $htmlsb,
			'Type',                     $htmlsb,
			'Special Poll Interval',    $htmlsb,
			'VMIC Time Provider',       $htmlsb
		)
#wv "columnWidths count $( $columnWidths.Count )" #MBS
#wv "columnHeaders count $( $columnHeaders.Count )" #MBS

		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '500'
		WriteHTMLLine 0 0 ''
	}

	Write-Verbose "$(Get-Date -Format G): Finished Create Domain Controller Time Server Configuration"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region EventLogInfo
Function ProcessEventLogInfo
{
	#Domain Controller Event Log Data
	$ContinueOn = $True
	If($MSWord -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Domain Controller Event Log Data"
		If(! ($Script:DARights -and $Script:Elevated) )
		{
			WriteWordLine 0 0 "To obtain event log data, you must run this script elevated and from an account with Domain Admin rights."
			$ContinueOn = $False
		}
	}
	If($Text)
	{
		Line 0 "///  Domain Controller Event Log Data  \\\"
		Line 0 ""
		If(! ($Script:DARights -and $Script:Elevated) )
		{
			Line 0 "To obtain event log data, you must run this script elevated and from an account with Domain Admin rights."
			$ContinueOn = $False
		}
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Controller Event Log Data&nbsp;&nbsp;\\\"
		If(! ($Script:DARights -and $Script:Elevated) )
		{
			WriteHTMLLine 0 0 "To obtain event log data, you must run this script elevated and from an account with Domain Admin rights."
			$ContinueOn = $False
		}
	}
	
	If($ContinueOn -eq $False)
	{
		Return
	}
	
	Write-Verbose "$(Get-Date -Format G): Create Domain Controller Event Log Data"
	Write-Verbose "$(Get-Date -Format G): `tAdd Domain Controller Event Log Data table to doc"
	
	#sort by DC and then event log name
	$xEventLogInfo = @($Script:DCEventLogInfo | Sort-Object EventLogName, DCName)
	
	If($MSWord -or $PDF)
	{
		$ELWordTable = New-Object System.Collections.ArrayList
	}
	If($HTML)
	{
		#V3.00 - pre-allocate
		#$rowdata = @()
#wv "Domain Controller Event Log Data"
#wv "rowCt $( $xEventLogInfo.Count )" #MBS
		$rowData = New-Object System.Array[] $xEventLogInfo.Count
		$rowIndx = 0
	}

	ForEach($Item in $xEventLogInfo)
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			EventLogName = $Item.EventLogName; 
			DCName = $Item.DCName; 
			EventLogSize = $Item.EventLogSize
			}

			## Add the hash to the array
			$ELWordTable.Add($WordTableRowHash) > $Null
		}
		If($Text)
		{
			Line 1 "Event Log Name`t`t: " $Item.EventLogName
			Line 1 "DC Name`t`t`t: " $Item.DCName
			Line 1 "Event Log Size (KB)`t: " $Item.EventLogSize
			Line 0 ""
		}
		If($HTML)
		{
#wv "rowIndx $rowIndx EventLogName $( $Item.EventLogName ) DCName $( $Item.DCName )" #MBS
			$rowdata[ $rowIndx ] = @(
				$Item.EventLogName, $htmlwhite,
				$Item.DCName,       $htmlwhite,
				$Item.EventLogSize, ( $htmlwhite -bor $global:htmlAlignRight ) #3.06 add alignright
			)
			$rowIndx++
		}
	}

	If($MSWord -or $PDF)
	{
		If($ELWordTable.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $ELWordTable `
			-Columns EventLogName, DCName, EventLogSize `
			-Headers "Event Log Name", "DC Name", "Event Log Size (KB)" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			## IB - set column widths without recursion
			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 150;
			$Table.Columns.Item(3).Width = 100;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($Text)
	{
		#nothing to do
	}
	If($HTML)
	{
		$columnWidths  = @( '300px', '150px', '90px' )
		$columnHeaders = @(
			'Event Log Name',      $htmlsb,
			'DC Name',             $htmlsb,
			'Event Log Size (KB)', $htmlsb
		)
		#wv "columnWidths count $( $columnWidths.Count )" #MBS
		#wv "columnHeaders count $( $columnHeaders.Count )" #MBS
		
		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '550'
		WriteHTMLLine 0 0 ''
	}

	Write-Verbose "$(Get-Date -Format G): Finished Create Domain Controller Event Log Data"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region SYSVOLStateInfo
Function ProcessSYSVOLStateInfo
{
	#new function added in 3.06
	#Domain Controller SYSVOL State Data
	If($MSWord -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Domain Controller SYSVOL State Data"
	}
	If($Text)
	{
		Line 0 "///  Domain Controller SYSVOL State Data  \\\"
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Controller SYSVOL State Data&nbsp;&nbsp;\\\"
	}
	
	Write-Verbose "$(Get-Date -Format G): Create Domain Controller SYSVOL State Data"
	Write-Verbose "$(Get-Date -Format G): `tAdd Domain Controller SYSVOL State Data table to doc"
	
	#sort by DC
	$xSYSVOLStateInfo = @($Script:DCSYSVOLState | Sort-Object DCName)
	
	If($MSWord -or $PDF)
	{
		$SSWordTable = New-Object System.Collections.ArrayList
		$WordHighlightedCells = New-Object System.Collections.ArrayList
		[int] $CurrentServiceIndex = 1
	}
	If($HTML)
	{
		$rowData = New-Object System.Array[] $xSYSVOLStateInfo.Count
		$rowIndx = 0
	}

	ForEach($Item in $xSYSVOLStateInfo)
	{
		If($MSWord -or $PDF)
		{
			$CurrentServiceIndex++
			
			$WordTableRowHash = @{ 
			DCName      = $Item.DCName; 
			SYSVOLState = $Item.SYSVOLState
			}

			## Add the hash to the array
			$SSWordTable.Add($WordTableRowHash) > $Null

			If($Item.SYSVOLState -NotLike "4 = Normal")
			{
				$WordHighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 2; }) > $Null
			}
		}
		If($Text)
		{
			Line 1 "DC Name`t`t: " $Item.DCName
			If($Item.SYSVOLState -NotLike "4 = Normal")
			{
				Line 1 "SYSVOL State`t: " "***$($Item.SYSVOLState)***"
			}
			Else
			{
				Line 1 "SYSVOL State`t: " $Item.SYSVOLState
			}
			Line 0 ""
		}
		If($HTML)
		{
			If($Item.SYSVOLState -NotLike "4 = Normal")
			{
				$HTMLHighlightedCells = $htmlred
			}
			Else
			{
				$HTMLHighlightedCells = $htmlwhite
			} 
			
			$rowdata[ $rowIndx ] = @(
				$Item.DCName,      $htmlwhite,
				$Item.SYSVOLState, $HTMLHighlightedCells
			)
			$rowIndx++
		}
	}

	If($MSWord -or $PDF)
	{
		If($SSWordTable.Count -gt 0)	#3.07
		{
			$Table = AddWordTable -Hashtable $SSWordTable `
			-Columns DCName, SYSVOLState `
			-Headers "DC Name", "SYSVOL State" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
			If($WordHighlightedCells.Count -gt 0)
			{
				SetWordCellFormat -Coordinates $WordHighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;
			}

			## IB - set column widths without recursion
			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 200;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($Text)
	{
		#nothing to do
	}
	If($HTML)
	{
		$columnWidths  = @( '200px', '200px' )
		$columnHeaders = @(
			'DC Name',      $htmlsb,
			'SYSVOL State', $htmlsb
		)
		
		FormatHTMLTable -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth '400'
		WriteHTMLLine 0 0 ''
	}

	Write-Verbose "$(Get-Date -Format G): Finished Create Domain Controller SYSVOL State Data"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region general script Functions
Function OutputReportFooter
{
	#Added in 3.07
	<#
	Report Footer
		Report information:
			Created with: <Script Name> - Release Date: <Script Release Date>
			Script version: <Script Version>
			Started on <Date Time in Local Format>
			Elapsed time: nn days, nn hours, nn minutes, nn.nn seconds
			Ran from domain <Domain Name> by user <Username>
			Ran from the folder <Folder Name>

	Script Name and Script Release date are script-specific variables.
	Script version is a script variable.
	Start Date Time in Local Format is a script variable.
	Domain Name is $env:USERDNSDOMAIN.
	Username is $env:USERNAME.
	Folder Name is a script variable.
	#>

	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Report Footer"
		WriteWordLine 2 0 "Report Information:"
		WriteWordLine 0 1 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		WriteWordLine 0 1 "Script version: $Script:MyVersion"
		WriteWordLine 0 1 "Started on $Script:StartTime"
		WriteWordLine 0 1 "Elapsed time: $Str"
		WriteWordLine 0 1 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		WriteWordLine 0 1 "Ran from the folder $Script:pwdpath"
	}
	If($Text)
	{
		Line 0 "///  Report Footer  \\\"
		Line 1 "Report Information:"
		Line 2 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		Line 2 "Script version: $Script:MyVersion"
		Line 2 "Started on $Script:StartTime"
		Line 2 "Elapsed time: $Str"
		Line 2 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		Line 2 "Ran from the folder $Script:pwdpath"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Report Footer&nbsp;&nbsp;\\\"
		WriteHTMLLine 2 0 "Report Information:"
		WriteHTMLLine 0 1 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		WriteHTMLLine 0 1 "Script version: $Script:MyVersion"
		WriteHTMLLine 0 1 "Started on $Script:StartTime"
		WriteHTMLLine 0 1 "Elapsed time: $Str"
		WriteHTMLLine 0 1 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		WriteHTMLLine 0 1 "Ran from the folder $Script:pwdpath"
	}
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	If($Text)
	{
		SaveandCloseTextDocument
	}
	If($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($MSWord)
	{
		If(Test-Path "$($Script:WordFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:WordFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:WordFileName)"
			Write-Error "Unable to save the output file, $($Script:WordFileName)"
		}
	}
	If($PDF)
	{
		If(Test-Path "$($Script:PDFFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:PDFFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:PDFFileName)"
			Write-Error "Unable to save the output file, $($Script:PDFFileName)"
		}
	}
	If($Text)
	{
		If(Test-Path "$($Script:TextFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:TextFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:TextFileName)"
			Write-Error "Unable to save the output file, $($Script:TextFileName)"
		}
	}
	If($HTML)
	{
		If(Test-Path "$($Script:HTMLFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:HTMLFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:HTMLFileName)"
			Write-Error "Unable to save the output file, $($Script:HTMLFileName)"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		$emailattachments = @()
		If($MSWord)
		{
			$emailAttachments += $Script:WordFileName
		}
		If($PDF)
		{
			$emailAttachments += $Script:PDFFileName
		}
		If($Text)
		{
			$emailAttachments += $Script:TextFileName
		}
		If($HTML)
		{
			$emailAttachments += $Script:HTMLFileName
		}
		SendEmail $emailAttachments
	}
}
#endregion

#region script start Function
Function ProcessScriptStart
{
	$script:startTime = Get-Date
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date -Format G): Script has completed"
	Write-Verbose "$(Get-Date -Format G): "

	Write-Verbose "$(Get-Date -Format G): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date -Format G): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date -Format G): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$Script:pwdpath\ADInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime   : $AddDateTime" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Company Name   : $Script:CoName" 4>$Null	#3.06 always output
		Out-File -FilePath $SIFile -Append -InputObject "ComputerName   : $ComputerName" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Address: $CompanyAddress" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email  : $CompanyEmail" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax    : $CompanyFax" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone  : $CompanyPhone" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page     : $CoverPage" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "DCDNSInfo      : $DCDNSInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev            : $Dev" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile   : $Script:DevErrorFile" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Domain name    : $ADDomain" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elevated       : $Script:Elevated" 4>$Null
		If($MSWord)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word FileName  : $($Script:WordFileName)" 4>$Null
		}
		If($HTML)
		{
			Out-File -FilePath $SIFile -Append -InputObject "HTML FileName  : $($Script:HTMLFileName)" 4>$Null
		}
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "PDF Filename   : $($Script:PDFFileName)" 4>$Null
		}
		If($Text)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Text FileName  : $($Script:TextFileName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder         : $Folder" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Forest name    : $ADForest" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From           : $From" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "GPOInheritance : $GPOInheritance" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "HW Inventory   : $Hardware" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "IncludeUserInfo: $IncludeUserInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log            : $($Log)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "MaxDetails     : $MaxDetails" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Report Footer  : $ReportFooter" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML   : $HTML" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF    : $PDF" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT   : $TEXT" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD   : $MSWORD" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info    : $ScriptInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Services       : $Services" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port      : $SmtpPort" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server    : $SmtpServer" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title          : $Script:Title" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To             : $To" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL        : $UseSSL" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name      : $UserName" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected    : $Script:RunningOS" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version   : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture      : $PSCulture" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture    : $PSUICulture" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language  : $Script:WordLanguageValue" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version   : $Script:WordProduct" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start   : $Script:StartTime" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time   : $Str" 4>$Null
	}

	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date -Format G): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date -Format G): Transcript/log stop failed"
			}
		}
	}
	#$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region do Collect()
#V3.00 - get timing on how long a [GC]::Collect() takes
Function ProcessGCCollect
{
	Param
	(
		[String] $tag
	)

	#Write-Verbose "$(Get-Date -Format G): Begin [GC]::Collect, tag = '$tag'"
	[System.GC]::Collect()
	#Write-Verbose "$(Get-Date -Format G): End [GC]::Collect"
}
#endregion

#region script core
#Script begins

ProcessScriptStart

ProcessScriptSetup

If($ADDomain -ne "")
{
	SetFilenames "$Script:DomainDNSRoot"
}
Else
{
	SetFilenames "$Script:ForestRootDomain"
}

If($Section -eq "All" -or $Section -eq "Forest")
{
	ProcessForestInformation

	ProcessAllDCsInTheForest
	
	ProcessCAInformation
	
	ProcessADOptionalFeatures
	
	ProcessADSchemaItems

	ProcessGCCollect 'Forest'
}

If($Section -eq "All" -or $Section -eq "Sites")
{
	ProcessSiteInformation
	ProcessGCCollect 'Sites'
}

If($Section -eq "All" -or $Section -eq "Domains")
{
	ProcessDomains
	ProcessDomainControllers
	ProcessGCCollect 'Domains-1'
}

If($Section -eq "All" -or $Section -eq "OUs")
{
	ProcessOrganizationalUnits
	ProcessGCCollect 'OUs'
}

If($Section -eq "All" -or $Section -eq "Groups")
{
	ProcessGroupInformation
	ProcessGCCollect 'Groups'
}

If($Section -eq "All" -or $Section -eq "GPOs")
{
	ProcessGPOsByDomain

	If($GPOInheritance -eq $True)
	{
		ProcessgGPOsByOUNew
	}
	Else
	{
		ProcessgGPOsByOUOld
	}
	
	ProcessOUsForBlockedInheritance
	
	ProcessGCCollect 'GPOs'
}

If($Section -eq "All" -or $Section -eq "Misc")
{
	ProcessMiscDataByDomain
	ProcessGCCollect 'Misc'
}

If(($Section -eq "All" -or $Section -eq "Domains"))
{
	#V3.00 combined these three into one "If"
	If($DCDnsInfo) #V3.04, only process if DCDNSInfo is true. No need for an empty table and Appendix otherwise
	{
		ProcessDCDNSInfo
	}
	ProcessTimeServerInfo
	ProcessEventLogInfo
	ProcessSYSVOLStateInfo
	ProcessGCCollect 'Domains-2'
}
#endregion

#region finish script
Write-Verbose "$(Get-Date -Format G): Finishing up document"
#end of document processing

If(($MSWORD -or $PDF) -and ($Script:CoverPagesExist))
{
	$AbstractTitle = "Microsoft Active Directory Inventory Report $script:MyVersion"
	$SubjectTitle = "Active Directory Inventory Report $script:MyVersion"
	UpdateDocumentProperties $AbstractTitle $SubjectTitle
}

If($ReportFooter)
{
	OutputReportFooter
}

ProcessDocumentOutput

ProcessScriptEnd
ProcessGCCollect 'ScriptEnd'
#endregion
