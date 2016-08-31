[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')
#https://en.wikipedia.org/wiki/List_of_Microsoft_Office_filename_extensions
function Expand-OOXMLDocument {
    param (
        [System.IO.FileInfo]$Document,
        [System.IO.DirectoryInfo]$Destination
    )
    [System.IO.Compression.ZipFile]::ExtractToDirectory($Document.FullName, $Destination.FullName)
}

function Compress-FolderIntoOOXMLDocument {
    param (
        [System.IO.DirectoryInfo]$DirectoryContainingSourceControlledVisioDocument,
        [System.IO.FileInfo]$Document
    )
    [System.IO.Compression.ZipFile]::CreateFromDirectory($DirectoryContainingSourceControlledVisioDocument.FullName, $Document.FullName)
}

function Edit-SourceControlledOOXMLDocument {
    param (
        [System.IO.DirectoryInfo]$DirectoryContainingSourceControlleOOXMLDocument
    )
    try {
        [System.IO.FileInfo]$DocumentToCreate = [System.IO.FileInfo]"$($DirectoryContainingSourceControlledVisioDocument.FullName).vsdx"
        Compress-FolderIntoVisioDocument -DirectoryContainingSourceControlledVisioDocument $DirectoryContainingSourceControlledVisioDocument -Document $DocumentToCreate  
        Start-Process "C:\Program Files (x86)\Microsoft Office\Office15\VISIO.EXE" -NoNewWindow -Wait -ArgumentList $DocumentToCreate.FullName
        Remove-Item -Recurse $DirectoryContainingSourceControlledVisioDocument
        expand-VisioDocument -Document $DocumentToCreate -Destination $DirectoryContainingSourceControlledVisioDocument
        Remove-Item "$($DirectoryContainingSourceControlledVisioDocument.FullName).vsdx"
        
        $XmlFiles = gci -Recurse $DirectoryContainingSourceControlledVisioDocument | 
        where extension -eq ".xml"

        $XmlFiles | %{
            $_;
            [xml](gc $_.FullName) | Format-XML | Set-Content -Encoding UTF8 -path $_.fullname
        }
    } catch {
        if (Test-Path -Path "$($DirectoryContainingSourceControlledVisioDocument.FullName).vsdx") {
            Remove-Item "$($DirectoryContainingSourceControlledVisioDocument.FullName).vsdx"
        }

        $_.Exception|format-list -force
    }
}

#http://blogs.msdn.com/b/powershell/archive/2008/01/18/format-xml.aspx
function Format-XML {
    param(
        [Parameter(
            Position=0, 
            Mandatory=$true, 
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )][xml]$xml, 
        $indent=2
    )
    process {
        $StringWriter = New-Object System.IO.StringWriter 
        $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
        $xmlWriter.Formatting = "indented" 
        $xmlWriter.Indentation = $Indent 
        $xml.WriteContentTo($XmlWriter) 
        $XmlWriter.Flush() 
        $StringWriter.Flush() 
        Write-Output $StringWriter.ToString() 
    }
}

# Copyright (c) 2014 Atif Aziz. All rights reserved.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#    http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

function Export-ExcelProject
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory = $true, 
                   HelpMessage = 'Specifies the path to the Excel Workbook file')]
        [string]$WorkbookPath,
        [Parameter(HelpMessage = 'Specifies export directory')]
        [string]$OutputPath,
        [Parameter(HelpMessage = 'Regular expression pattern identifying modules to be excluded')]
        [string]$Exclude,
        [Parameter(HelpMessage = 'Export items that may be auto-named, like Class1, Module2, etc.')]
        [switch]$IncludeAutoNamed = $false,
        [switch]$Force = $false
    )
    
    function Get-MD5Hash($filePath)
    { 
        $bytes = [IO.File]::ReadAllBytes($filePath)
        $hash = [Security.Cryptography.MD5]::Create().ComputeHash($bytes)
        [BitConverter]::ToString($hash).Replace('-', '').ToLowerInvariant()
    }
    
    $mo = Get-ItemProperty -Path HKCU:Software\Microsoft\Office\*\Excel\Security `
                           -Name AccessVBOM `
                           -EA SilentlyContinue | `
              ? { !($_.AccessVBOM -eq 0) } | `
              Measure-Object

    if ($mo.Count -eq 0)
    {
        Write-Warning 'Access to VBA project model may be denied due to security configuration.'
    }

    Write-Verbose 'Starting Excel'
    $xl = New-Object -ComObject Excel.Application -EA Stop
    Write-Verbose "Excel $($xl.Version) started"
    $xl.DisplayAlerts = $false
    $missing = [Type]::Missing
    $extByComponentType =  @{ 100 = '.cls'; 1 = '.bas'; 2 = '.cls' }
    $outputPath = ($outputPath, (Get-Item .).FullName)[[String]::IsNullOrEmpty($outputPath)]
    mkdir -EA Stop -Force $outputPath | Out-Null
    
    try
    {
        # Open(Filename, [UpdateLinks], [ReadOnly], [Format], [Password], [WriteResPassword], [IgnoreReadOnlyRecommended], [Origin], [Delimiter], [Editable], [Notify], [Converter], [AddToMru], [Local], [CorruptLoad]) 
        $wb = $xl.Workbooks.Open($workbookPath, $false, $true, `
                                 $missing, $missing, $missing, $missing, $missing, $missing, $missing, $missing, $missing, $missing, $missing, `
                                 $true)
        
        $wb | Get-Member | Out-Null # HACK! Don't know why but next line doesn't work without this
        $project = $wb.VBProject
        
        if ($project -eq $null)
        {
            Write-Verbose 'No VBA project found in workbook'
        }
        else
        {
            $tempFilePath = [IO.Path]::GetTempFileName()

            $vbcomps = $project.VBComponents
            
            if (![String]::IsNullOrEmpty($exclude))
            {
                $verbose = ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose') -and $PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent)
                if ($verbose) 
                {
                    $vbcomps | ? { $_.Name -match $exclude } | % { Write-Verbose "$($_.Name) will be excluded" }
                }
                $vbcomps = $vbcomps | ? { $_.Name -notmatch $exclude }
            }
            
            $vbcomps | % `
            { 
                $vbcomp = $_
                $name = $vbcomp.Name
                $ext = $extByComponentType[$vbcomp.Type]
                if ($ext -eq $null)
                {
                    Write-Verbose "Skipped component: $($name)"
                }
                elseif (!$includeAutoNamed -and $name -match '^(Form|Module|Class|Sheet)[0-9]+$')
                {
                    Write-Verbose "Skipped possibly auto-named component: $name"
                }
                else
                {
                    $vbcomp.Export($tempFilePath)
                    
                    $exportedFilePath = Join-Path $outputPath "$name$ext"
                    $exists = Test-Path $exportedFilePath -PathType Leaf
                    
                    if ($exists) 
                    { 
                        $oldHash = Get-MD5Hash $exportedFilePath 
                        $newHash = Get-MD5Hash $tempFilePath
                        $changed = !($oldHash -eq $newHash)
                        $status  = ('Unchanged', 'Conflict', 'Unchanged', 'Changed')[[int]$changed + (2 * [int]$force.IsPresent)]
                    }
                    else
                    {
                        $status = 'New'
                    }

                    if (($status -eq 'Changed' -or $status -eq 'New') `
                        -and $pscmdlet.ShouldProcess($name))
                    {
                        Move-Item -Force $tempFilePath $exportedFilePath
                    }
                    
                    New-Object PSObject -Property @{
                        Name   = $name;
                        Status = $status;
                        File   = (Get-Item $exportedFilePath -EA Stop);
                    }
                }
            }        
        }
        $wb.Close($false, $missing, $missing)
    }
    finally
    {    
        $xl.Quit()
        # http://technet.microsoft.com/en-us/library/ff730962.aspx
        [Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$xl) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}