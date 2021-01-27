<# 	PowerShell script to remove timestamp in Word's track changes and comments.
	You can do this manually my renaming the Word.docx file to file.zip and open this Archive.
	The Zip-Archive will have several folders. Change to the "word" folder, open the files document.xml, 
	comments.xml and footnotes.xml.
	You can open this files with any Texteditor. Search and replace
	w:author="<YourSurename, YourName> (<YourOrganization>)" w:date="2020-12-11T08:18:00Z 
	with this search and replace setting:
	search: w:author="<YourSurename, YourName> (>YourOrganization>)" w:date
	replace: w:author="<YourSurename, YourName> (>YourOrganization>)" w:ignore
	It will keep the author name in place and hide the timestamp in Word.
	
	But this solution will not remove the timestamp from the three files. It will hide the timestamp 
	in Word only. This might good enough for entry level users but not for experts like you.
	Some Texteditors like Notepad++	or Geany let you search/replace regex. You can change the timestamp 
	as regex with it. In some enviroments you don't have this tools by hand and you are not able to install Software. 
	You can change the timestamp by hand, one by one because every timestamp is different.
	
	In this case, when you work with restricted access, you can do the task with a PowerShell script.
	I first tried to use a batch file. But it did not work for me. 
	
	Move the Word docx file into the same folder this script is stored in. 
	Open a Powershell Terminal and set the environment by entering this command: Set-ExecutionPolicy RemoteSigned -force
	Start the script by entering .\Word-remove-timestamp.ps1 (or whatever). When you run the script,
	it will ask you for the specific Author name and Organization name you want to free from the timestamp.
	It then will show you the docx files in the folder and let you choose the file. It renames the file you choose
	to a .zip file and extract it.
	It starts to replace the word "date" with "ignore". It also replace the timestamp with the static date 1970-01-01T00:00:00Z.
	
	After that, it stores the files back to the folder, zip the folders and rename the .zip file back to .docx.
	
	I know there could be more code maybe for checking if the file or folder exist and so on. But for me this
	script works. Feel free to add more functions to it. 
#>

# Set-ExecutionPolicy RemoteSigned -force
# Write-Host "The Author's Surename, Name and in the next step the Organization name is required."

param (
	[Parameter(Mandatory)]
    $Author,
    
    [Parameter(Mandatory)]
    $Organization
)


# Let's see what .docx files we have in this folder. 
Get-ChildItem -Path .\*.docx -Force
$WordDoc = Read-Host -Prompt 'Enter the filename you want to change (without extension .docx).'
# Enter your Surename, Name.
#$Author = Read-Host -Prompt 'Enter your Surename, Name'
# Enter your Organization name (maybe check the document.xml file first).
#$Organization = Read-Host -Prompt 'Enter your Organisation short'

# We will search for this pattern in the files.


#$Author = [Regex]::Escape($Author)
$Organization = [Regex]::Escape($Organization)

$Ext = '.docx'
$WordDocExt = $WordDoc+$Ext
$New = "-2"
$WordDocNew = $WordDoc+$New
$WordDocNewExt = $WordDocNew+$Ext
$Zip = ".zip"
$ZipDoc = $WordDoc+$Zip
$ZipDocNew = $WordDocNew+$Zip


# Let's copy the .docx to .zip
Copy-Item $($WordDocExt) -Destination $($ZipDoc)

Expand-Archive -Path .\$($ZipDoc) -DestinationPath .\$($WordDoc)

# We don't need the Zip file anymore. Let's delete it.
Remove-Item $($ZipDoc)

#
# We are looking for 	w:author="Surename, Name (Organization)" w:date="2020-12-11T08:18:00Z"
# and want to replace                                     w:ignore="1970-01-01T00:00:00Z"
# where $Author is param {0} and $Organization is param {1}
#
<#
# Let's replace the timestamp in documents.xml. 
(Get-Content .\$($WordDoc)\word\document.xml) -replace @(
	 'w:author="[^"]+" w:date="[^"]+"'
    'w:author="{0} ({1})" w:ignore="1970-01-01T00:00:00Z"' -f $Author, $Organization
) | Set-Content .\$($WordDoc)\word\document.xml

# Let's replace the timestamp in comments.xml. 
(Get-Content .\$($WordDoc)\word\comments.xml) -replace @(
	 'w:author="[^"]+" w:date="[^"]+"'
    'w:author="{0} ({1})" w:ignore="1970-01-01T00:00:00Z"' -f $Author, $Organization
) | Set-Content .\$($WordDoc)\word\comments.xml
#>


# Let's replace the timestamp in document.xml. 
(Get-Content .\$($WordDoc)\word\document.xml) -replace @(
	       '(?<=w:author="{0} ({1})" )w:date="[^"]+"' -f $Author, $Organization
           'w:ignore="1970-01-01T00:00:00Z"'
) | Set-Content .\$($WordDoc)\word\document.xml


# Let's replace the timestamp in comments.xml. 
(Get-Content .\$($WordDoc)\word\comments.xml) -replace @(
	        '(?<=w:author="{0} ({1})" )w:date="[^"]+"' -f $Author, $Organization
            'w:ignore="1970-01-01T00:00:00Z"'
) | Set-Content .\$($WordDoc)\word\comments.xml

# Let's replace the timestamp in footnotes.xml. 
(Get-Content .\$($WordDoc)\word\footnotes.xml) -replace @(
	        '(?<=w:author="{0} ({1})" )w:date="[^"]+"' -f $Author, $Organization
            'w:ignore="1970-01-01T00:00:00Z"'
) | Set-Content .\$($WordDoc)\word\footnotes.xml


# Let's recreate the Zipfile
# 7zip a '$ZipDocNew' $WordDoc 
Compress-Archive -Path .\$($WordDoc)\* -CompressionLevel Optimal -DestinationPath .\$($ZipDocNew) 

# Let's rename the .zip back to .docx
Rename-Item -Path $($ZipDocNew) -NewName $($WordDocNewExt)
#Copy-Item $($ZipDocNew) -Destination $($WordDocNewExt)

# We don't need the Zip file anymore. Let's delete it.
Remove-Item -Recurse -Force $($WordDoc) 
# Remove-Item $($ZipDocNew)

