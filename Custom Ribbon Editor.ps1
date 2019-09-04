#add content to the customUI.xml
$XML = cat -raw ".\customUI\customUI.xml"

add-Type -AssemblyName System.IO.Compression.FileSystem
#this function unzips the $zipfile to the $outpath
function Unzip([string]$zipfile, [string]$outpath) {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

#this function zips $file to $outpath
function Zip([string]$file, [string]$outpath) {
    [io.compression.zipfile]::CreateFromDirectory($file, $outpath)
}

#this function adds attribute $Name with value $Value to $Node
function Add-XMLAttribute([System.Xml.XmlNode] $Node, $Name, $Value) {
  $attrib = $Node.OwnerDocument.CreateAttribute($Name)
  $attrib.Value = $Value
  $node.Attributes.Append($attrib)
}

#used to keep track of the file extension for conversion
$extension = ""

#the prompt to choose a file
write-output "Please choose a macro-enabled file."
while ($true) {
    #open the file explorer to get the user input for an excel file
    Add-Type -AssemblyName System.Windows.Forms
    $f = new-object Windows.Forms.OpenFileDialog
    $f.InitialDirectory = get-location
    $f.Filter = "All Files (*.*)|*.*"
    $f.ShowHelp = $true
    $f.Multiselect = $false
    [void]$f.ShowDialog()

    #user cancelled or exited the file explorer
    if ($f.FileName -eq "") {
        exit
    
    #use chose macro-enabled workbooks
    } elseif ((get-item $f.FileName).Extension -eq ".xlsm") {
        $extension = ".xlsm"
        write-output "Macro-enabled excel workbook(xlsm) chosen"
        Start-Sleep 1
        break
    } elseif ((get-item $f.FileName).Extension -eq ".docm") {
        $extension = ".docm"
        write-output "Macro-enabled word document(docm) chosen"
        Start-Sleep 1
        break
    } elseif ((get-item $f.FileName).Extension -eq ".pptm") {
        $extension = ".pptm"
        write-output "Macro-enabled powerpoint presentation(pptm) chosen"
        Start-Sleep 1
        break
    }
    write-output "Invalid file format chosen.  File must be macro-enabled.  Please choose again."
    Start-Sleep 2
}

#now we change the extension to a zip file
$fileName = (get-item $f.FileName).basename
$filePath = (get-item $f.FileName).DirectoryName + "\" + $fileName
$newName = $filePath + ".zip"

#change the extension of the file and check if a zip file already exists
if (Test-Path $newName) {
    Write-Output $newName + " already exists.  [Press any key to quit]."
    cmd /c pause | out-null
    exit
}

rename-item -path $f.FileName -newname $newName

#unzip the file to access the contents
try {
    Unzip $newName $filePath
    remove-item $newName -recurse
} catch {
    rename-item -path $newName -newname $f.FileName
    write-output "Your file seems to be corrupted.  No Changes Made.  [Press any key to quit]."
    cmd /c pause | out-null
    exit
}

#TODO
# change the way the file name is found so that other file names can be used instead of just "customUI14.xml"


$newXMLPath
$newXMLFolder = $filePath + "\customUI"

$newFolder = ""



#if the ribbon doesn't exist, ask if they want to add one
if (-not (Test-Path $newXMLFolder)) {
    while ($true) {
        $answer = read-host "There is currently no custom ribbon.  Would you like to add one? [Y/N]"
        #create a new custom ribbon
        if ($answer -eq "Y" -or $answer -eq "y") {
            $newXMLPath = $filePath + "\customUI\customUI.xml"
            write-output "You can customize the XML file to add personalized macros."
            #create a new folder to contain the changed XML
            $newFolder = (get-item $filePath).DirectoryName + "\customUI"
            New-Item -ItemType directory -Path $newFolder | out-null

            #add the XML to the new folder
            New-Item -Path $newFolder -Name "customUI.xml" -ItemType "file" -Value $XML| out-null

            #add the folder to the unzipped folder
            Copy-Item $newFolder $filePath -Recurse

            #remove the customUI Folder
            Remove-Item $newFolder -Recurse
            
            #update .rels file
            $relFile = $fileName + "\_rels\.rels"

            #initialize a new XMl object and add the necessary additional elements
            $doc = New-Object System.Xml.XmlDocument
            $doc.Load($relFile)
            $nameSpace = $doc.Relationships.NamespaceURI
            $child = $doc.CreateElement("Relationship", $nameSpace)
            Add-XMLAttribute $child "Target" "customUI/customUI.xml" | Out-Null
            Add-XMLAttribute $child "Type" "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" | Out-Null
            Add-XMLAttribute $child "Id" "R62d869510554412e" | Out-Null
            $doc.Relationships.AppendChild($child) | Out-Null
            $doc.Save($relFile)
            Start-Sleep 1
            Write-Output "Custom ribbon created."
            break
        #terminate the script
        } elseif ($answer -eq "N" -or $answer -eq "n") {
            Zip $filePath $newName
            
            #remove the uncompressed excel directory
            Remove-Item $filePath -Force -Confirm:$false -Recurse

            #change the file to a original file
            $oldName = (get-item $newName).DirectoryName + "\" + $fileName + $extension
            rename-item -path $newName -newname $oldName
            write-output "No changes were made.  [Press any key to quit]."
            cmd /c pause | out-null
            exit
        }
        write-output "Invalid command, please try again"
        Start-Sleep 1
    }
} else {
    Start-Sleep 1
    write-output "Custom ribbon found."
}

$newXMLPath = Get-ChildItem ($filePath + "\customUI\customUI*.xml")


Start-Sleep 1
write-output "Opening Notepad..."
notepad $newXMLPath
Start-Sleep 1

write-output "Save to keep your changes or quit without saving to discard them.  After making your changes, press any key to continue."
cmd /c pause | out-null

#zip the file
Zip $filePath $newName 

#remove the uncompressed excel directory
Remove-Item $filePath -Force -Confirm:$false -Recurse

#change the file to an xlsm
$oldName = (get-item $newName).DirectoryName + "\" + $fileName + $extension
rename-item -path $newName -newname $oldName

write-output "Custom ribbon updated.  [Press any key to exit]."
cmd /c pause | out-null
