#### loading assembly libaries
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")


#### add libaries
Add-Type -AssemblyName System
Add-Type -AssemblyName system.io.compression.filesystem


##### GLOBAL
function Global:FILESYSTEM() {
    return [System.IO.Compression.ZipFile]
} ### FILESYSTEM



##### datetime setting
class DialyDate_SETTING {
    ### private datas
    hidden $form
    hidden $okButton
    hidden $cancelButton
    hidden $pickDate


    ###public data
    [datetime]$dialyDate


    DialyDate_SETTING() {
        ### init setting
        $this.form = New-Object System.Windows.Forms.Form
        $this.okButton = New-Object System.Windows.Forms.Button
        $this.cancelButton = New-Object System.Windows.Forms.Button
        $this.pickDate = New-Object System.Windows.Forms.DateTimePicker

        $this.PickDate_SETTING()
        $this.OKBUtton_SETTING()
        $this.CancelButton_SETTING()
        $this.Form_SETTING()

        if ($this.okButton.Name -eq "OK") {
            $this.dialyDate = $this.pickDate.Value
        }
    }

    ### private function
    hidden Form_SETTING() {
        ##### main application setting
        $this.form.Width = $this.pickDate.Width + $this.okButton.Width + $this.cancelButton.Width
        $this.form.Height = $this.datetimePicker.Height
        $this.form.FormBorderStyle = 0
        $this.form.AutoSize = $true
        $this.form.MinimizeBox = $false
        $this.form.MaximizeBox = $false
        $this.form.StartPosition = 1
        ### application show dialog
        $this.form.ShowDialog()
    } ### main form setting
    hidden PickDate_SETTING() {
        $this.pickDate.Format = 8
        $this.pickDate.Font = New-Object System.Drawing.Font -ArgumentList 'arial', 20
        $this.pickDate.CustomFormat = "MMM-dd-yyyy"
        $this.pickDate.AutoSize = $true
        $this.pickDate.Name = 'date picker'
        $this.pickDate.Top = 3
        $this.form.Controls.Add($this.pickDate)
    } ### pick date setting
    hidden OKBUtton_SETTING() {
        $this.okButton.Text = "O"
        $this.okButton.Name = "CANCEL"
        $this.okButton.Font = $this.pickDate.Font
        $this.okButton.AutoSize = $true
        $this.okButton.BackColor = '#00FF00'
        $this.okButton.ForeColor = '#FFFFFF'
        $this.okButton.Add_Click({
            param($sender, $args)
            $sender.Name = "OK"
            $sender.Parent.Close()
        })
        $this.okButton.Left = $this.pickDate.Right
        $this.form.Controls.Add($this.okButton)
    } ### OK button setting
    hidden CancelButton_SETTING() {
        $this.cancelButton.Text = 'X'
        $this.cancelButton.BackColor = '#FF0000'
        $this.cancelButton.ForeColor = '#ffffff'
        $this.cancelButton.Left = $this.okButton.Right
        $this.cancelButton.Font = $this.pickDate.Font
        $this.cancelButton.AutoSize = $true
        $this.cancelButton.Add_Click({
            param($sender, $args)

            $sender.Parent.Close()
        })
        $this.form.Controls.Add($this.cancelButton)
    } ### cancel button setting
}

#### init program
class Program {
    ### private data
    hidden [string]$dateFormat = "ddMMMyyyy"
    hidden [string]$currentDir
    hidden [string]$tmpFolder

    Program() {
        $dialyDate_SETTING  = New-Object DialyDate_SETTING
        [datetime]$dialyDate = $dialyDate_SETTING.dialyDate
        $dialyDate = [datetime]::new($dialyDate.Year, $dialyDate.Month, $dialyDate.Day, 0, 0, 0)
        if ($dialyDate.Year -ge 2022) {
            ## setting init data
            $this.tmpFolder = "$(Get-Location)\tmpdata\$(Get-Date -Format 'ddMMyyyy_HHmmss')"
            

            ## check folder
            $this.Update_FILEANDFOLDER()

            ## read saved data
            $networks = (Get-Content -Path "$(Get-Location)/networks.non")
            [GettingFile]$gettingFiles = New-Object GettingFile -ArgumentList $networks, $dialyDate
            $this.Update($gettingFiles.dialySNs, $gettingFiles.tmp)

            foreach($file in $gettingFiles.tmp) {
                copy $file $this.tmpFolder
                Write-Host ([string]::Format("Copy file: {0}", $file))
            }
        }
    }

    ##### private functions
    hidden [void]Update_FILEANDFOLDER() {
        Clear-Content .\data.txt
        $tmpdataFolder = "$(Get-Location)/tmpdata"
        $networkFile = "$(Get-Location)/networks.non"
        if ((Test-Path -Path $networkFile -PathType Leaf) -eq $false) {
            Out-File -FilePath $networkFile
        }
        if ((Test-Path -Path $tmpdataFolder) -eq $false) {
            New-Item -ItemType Directory -Path $tmpdataFolder
        }
        New-Item -ItemType Directory -Path $this.tmpFolder
    }#### checking file and forder

    
    hidden Update($dialySNs, $failedFiles) {
        foreach($SN in $dialySNs) {
            $fileList = New-Object System.Collections.Generic.List[string]
            foreach($file in $failedFiles) {
                if ($file.Contains($SN)) {
                    $fileList.Add($file)
                }
            }
            [DataReader]$dataReader = New-Object DataReader -ArgumentList $SN, $fileList


            ### reading data after all checked
            $readData = "$($dataReader.data.model);$SN;$($dataReader.data.station);$($dataReader.data.tester);$($dataReader.data.pallet);$($dataReader.data.revison);$($dataReader.data.testTime);$($dataReader.data.firstFailed);$($dataReader.data.defectValue)"
            $readData | Add-Content .\data.txt
            Write-Host $dataReader.data.ext
            Write-Host $readData
        }
    }
}

##### getting file in networks
class GettingFile {
    ### private data
    hidden $allSNs
    hidden $allFailed
    hidden $files
    hidden $totalFile = 0


    ### public data
    $tmp
    $dialySNs


    GettingFile([string[]]$networks, [datetime]$dialyDate) {
        $this.allSNs = New-Object System.Collections.Generic.List[string]
        $this.allFailed = New-Object System.Collections.Generic.List[string]
        $this.dialySNs = New-Object System.Collections.Generic.List[string]
        $this.files = New-Object System.Collections.Generic.List[string]
        $this.tmp = New-Object System.Collections.Generic.List[string]
        
        foreach($network in $networks) {
            Write-Host ([string]::Format("Getting file from network: {0}", $network))
            if ((Test-Path -Path $network) -eq $true) {
                $fs = Get-ChildItem -Path $network -Recurse
                foreach($f in $fs) {
                    if ($f.Name.Length -ge 17) {
                        $this.files.Add($f.FullName)
                    }
                }
            }
        }
        Write-Host ([string]::Format("Total files: {0}", $this.files.Count))
        $this.totalFile = $this.files.Count
        $this.Sorting($dialyDate)
    }


    ### private function
    hidden Sorting([datetime]$dialyDate) {
        $bSN = New-Object System.Collections.Generic.List[string]
        foreach($file in $this.files) {
            if ((Test-Path -Path $file -PathType Leaf) -eq $true) {
                $getFile = (Get-Item -Path $file)
                $SN = $getFile.Name.Substring(0, 17)
                if ($this.CheckSN($SN) -eq $true) {
                    $fileDate = ""
                    if ($getFile.Extension -eq '.txt') {
                        if ((Get-Content -Path $file | Out-String).Contains("Test Result: FAIL") -eq $true) {
                            $this.allFailed.Add($file)
                            $fileDate = $getFile.LastWriteTime.ToString("ddMMMyyyy")
                            if ([datetime]::new($getFile.LastWriteTime.Year, $getFile.LastWriteTime.Month, $getFile.LastWriteTime.Day, $getFile.LastWriteTime.Hour, $getFile.LastWriteTime.Minute, $getFile.LastWriteTime.Second) -lt $dialyDate) {
                                if ($bSN.Contains($SN) -eq $false) {
                                    $bSN.Add($SN)
                                }   
                            }
                            
                        }
                    }
                    else {
                        if ($getFile.Name.Contains("fail") -or $getFile.Name.Contains("inconclusive")) {
                            $this.allFailed.Add($file)
                            if ($getFile.Extension -eq '.zip') {
                                $zip = (Global:FILESYSTEM)::OpenRead($file)
                                $zipFiles = $zip.Entries
                                foreach($zipFile in $zipFiles) {
                                    if ($zipFile.Name.Contains(".html") -and $zipFile.Name.Contains($SN)) {
                                        $fileDate = $zipFile.LastWriteTime.ToString("ddMMMyyyy")
                                        if ([datetime]::new($zipFile.LastWriteTime.Year, $zipFile.LastWriteTime.Month, $zipFile.LastWriteTime.Day, $zipFile.LastWriteTime.Hour, $zipFile.LastWriteTime.Minute, $zipFile.LastWriteTime.Second) -lt $dialyDate) {
                                            if ($bSN.Contains($SN) -eq $false) {
                                                $bSN.Add($SN)
                                            } 
                                        }
                                    }
                                }
                                $zip.Dispose()
                            }
                            else {
                                $fileDate = $getFile.LastWriteTime.ToString("ddMMMyyyy")
                                if ([datetime]::new($getFile.LastWriteTime.Year, $getFile.LastWriteTime.Month, $getFile.LastWriteTime.Day, $getFile.LastWriteTime.Hour, $getFile.LastWriteTime.Minute, $getFile.LastWriteTime.Second) -lt $dialyDate) {
                                    if ($bSN.Contains($SN) -eq $false) {
                                        $bSN.Add($SN)
                                    }   
                                }
                            }
                        }
                    }

                    if ($this.allSNs.Contains($SN) -eq $false -and $fileDate -eq $dialyDate.ToString("ddMMMyyyy")) {
                         $this.allSNs.Add($SN)     
                    }
                }
            }
            $this.totalFile--
            Write-Host ([string]::Format("Cheking file Remain: {0}", $this.totalFile))
        }
        foreach($SN in $this.allSNs) {
            if ($bSN.Contains($SN) -eq $false -and $this.dialySNs.Contains($SN) -eq $false) {
                $this.dialySNs.Add($SN)
            }
        }
        foreach($SN in $this.dialySNs) {
            foreach($file in $this.allFailed) {
                if ($file.Contains($SN) -eq $true) {
                    $this.tmp.Add($file)
                }
            }
        }
    }


    hidden [bool]CheckSN($SN) {
        $result = $false
        $workorder = $SN.Substring(2, 8)
        $number = $SN.Substring(11, 6)
        try {
            [convert]::ToInt32($workorder)
            [convert]::ToInt32($number)
            $result = $true
        }
        catch {}
        return $result
    } ### checking unit serie number
}


##### reading data
class DataReader {
    ### public data
    $data


    DataReader($SN, $fileList) {
        if ($fileList.Count -gt 0 -and [string]::IsNullOrEmpty($SN) -eq $false) {
            $lastFile = $this.GettingLastFile($fileList, $SN)
            if ([string]::IsNullOrEmpty($lastFile) -eq $false) {
                if ((Test-Path -Path $lastFile -PathType Leaf) -eq $true) {
                    $content = (Get-Content -Path $lastFile | Out-String)
                    if ((Get-Item -Path $lastFile).Extension -eq '.zip') {
                        $content = [string]::Empty
                        $zip = (Global:FILESYSTEM)::OpenRead($lastFile)
                        $zipFiles = $zip.Entries
                        foreach($zipFile in $zipFiles) {
                            if ($zipFile.Name.Contains(".html") -and $zipFile.Name.Contains($SN) -and [string]::IsNullOrEmpty($content) -eq $true) {
                                $stream = $zipFile.Open()
                                $reader = New-Object System.IO.StreamReader($stream)
                                $content = $reader.ReadToEnd()

                                $reader.Close()
                                $stream.Close()
                            }
                        }
                        $zip.Dispose

                        $this.data = New-Object DataReader_HTML -ArgumentList $content
                        $this.data.ext = '.zip'
                    }
                    elseif ((Get-Item -Path $lastFile).Extension -eq '.html') {
                        $this.data = New-Object DataReader_HTML -ArgumentList $content
                        $this.data.ext = '.html'
                    }
                    elseif ((Get-Item -Path $lastFile).Extension -eq '.xml') {
                        $this.data = New-Object DataReader_XML -ArgumentList $content
                        $this.data.ext = '.xml'
                    }
                    else {
                        $this.data = New-Object DataReader_TXT -ArgumentList $content
                        $this.data.ext = '.txt'
                    }


                    ##### clear invalid data
                    $this.data.model = $this.data.configName.Split("-")[0].Split("_")[0]
                    if ($this.data.ext -ne '.txt') {
                        if ($this.data.configName.Contains("PAT") -eq $true) {
                            $this.data.station = "ATS1"
                            $this.data.revison = $this.data.configName.Split("_")[0].Split("-")[$this.data.configName.Split("_")[0].Split("-").Length - 1]
                        }
                        elseif ($this.data.configName.Contains("FAT") -eq $true) {
                            $this.data.station = "ATS2"
                            $this.data.revison = $this.data.configName.Split("_")[0].Split("-")[$this.data.configName.Split("_")[0].Split("-").Length - 1]
                        }
                        elseif ($this.data.configName.Contains("ATS3") -eq $true) {
                            $this.data.station = "ATS3"
                        }
                        elseif ($this.data.configName.Contains("VehCAN") -eq $true) {
                            $this.data.station = "VFlash"
                        }
                        elseif ($this.data.configName.Contains("Std_BI") -eq $true) {
                            $this.data.station = "Heat up"
                        }
                    }
                    $this.data.testTime = (Get-Item -Path $lastFile).LastWriteTime.ToString("HH:mm")
                    $this.data.defectValue = $this.data.defectValue.Replace([System.Environment]::NewLine, "//")
                }
            }
        }
    }

    ### private function
    hidden [string]GettingLastFile($files, [string]$SN) {
        [string]$lastFile = [string]::Empty
        foreach($file in $files) {
            if ($file.Contains($SN)) {
                if ([string]::IsNullOrEmpty($lastFile)) {
                    $lastFile = $file
                }
                else {
                    $lastTime_00 = (Get-Item -Path $file).LastWriteTime
                    $lastTime_01 = (Get-Item -Path $lastFile).LastWriteTime
                    if ((Get-Item -Path $file).Extension -eq '.zip') {
                        $zip = (Global:FILESYSTEM)::OpenRead($file)
                        $zipFiles = $zip.Entries
                        foreach($zipFile in $zipFiles) {
                            if ($zipFile.Name.Contains(".html") -and $zipFile.Name.Contains($SN)) {
                                $lastTime_00 = $zipFile.LastWriteTime
                            }
                        }
                        $zip.Dispose()
                    }
                    if ((Get-Item -Path $lastFile).Extension -eq '.zip') {
                        $zip = (Global:FILESYSTEM)::OpenRead($lastFile)
                        $zipFiles = $zip.Entries
                        foreach($zipFile in $zipFiles) {
                            if ($zipFile.Name.Contains(".html") -and $zipFile.Name.Contains($SN)) {
                                $lastTime_01 = $zipFile.LastWriteTime
                            }
                        }
                        $zip.Dispose()
                    }
                    $lastTime_00 = [datetime]::new($lastTime_00.Year, $lastTime_00.Month, $lastTime_00.Day, $lastTime_00.Hour, $lastTime_00.Minute, $lastTime_00.Second)
                    $lastTime_01 = [datetime]::new($lastTime_01.Year, $lastTime_01.Month, $lastTime_01.Day, $lastTime_01.Hour, $lastTime_01.Minute, $lastTime_01.Second)


                    if ($lastTime_00 -gt $lastTime_01) {
                        $lastFile = $file
                    }
                }
            }
        }
        return $lastFile
    } ##### update last file
}

##### data reader xml
class DataReader_XML {
    ### private data
    hidden $xmlData = @{}


    ### public data
    [string]$configName = [string]::Empty
    [string]$ext = [string]::Empty
    [string]$model = [string]::Empty
    [string]$station = [string]::Empty
    [string]$tester = [string]::Empty
    [string]$pallet = [string]::Empty
    [string]$revison = [string]::Empty
    [string]$testTime = [string]::Empty
    [string]$firstFailed = [string]::Empty
    [string]$defectValue = [string]::Empty
    

    DataReader_XML([string]$content) {
        if ([string]::IsNullOrEmpty($content) -eq $false) {
            $this.GettingDocument($content)
            $this.Update()


            ### clear invalid data
            if (($this.station -eq "ATS1" -or $this.station -eq "ATS2") -eq $false) {
                $this.revison = [string]::Empty
            }
        }
    }


    ### private function
    hidden GettingDocument($content) {
        $this.xmlData.Add("xinfo", $this.GetXML_DATA($content, "xinfo"))
        $this.xmlData.Add("info", $this.GetXML_DATA($content, "info"))
        $this.xmlData.Add("testcase", $this.GetXML_DATA($content, "testcase"))
    } ### getting document
    hidden Update() {
        foreach($data in $this.xmlData["xinfo"]) {
            if ($data.Node.name -eq "Configuration" -and [string]::IsNullOrEmpty($this.configName)) {
                $this.configName = $data.Node.description.Split("\\")[$data.Node.description.Split("\\").Length - 1]
            }
            if ($data.Node.name -eq "Windows Computer Name" -and [string]::IsNullOrEmpty($this.tester)) {
                $this.tester = $data.Node.description
            }
        }
        foreach($data in $this.xmlData["info"]) {
            if ($data.Node.name -eq "Pallet_No" -and [string]::IsNullOrEmpty($this.pallet)) {
                $this.pallet = $data.Node.description
            }
        }
        foreach($data in $this.xmlData["testcase"]) {
            if (($data.Node.verdict.result -eq 'fail' -or $data.Node.verdict.result -eq 'inconclusive') -and [string]::IsNullOrEmpty($this.firstFailed)) {
                $this.firstFailed = $data.Node.title
                $this.defectValue = $data.Node.extendedinfo.InnerText
            }
        }
    } ### reading data
    hidden [object[]]GetXML_DATA($content, $XPath) {
        return (Select-Xml -Content $content -XPath ([string]::Format("//{0}", $XPath)))
    } ### data selection
}

##### data reader html
class DataReader_HTML {
    ### private data
    hidden $doc
    hidden $valueID = [string]::Empty


    ### public data
    [string]$configName = [string]::Empty
    [string]$ext = [string]::Empty
    [string]$model = [string]::Empty
    [string]$station = [string]::Empty
    [string]$tester = [string]::Empty
    [string]$pallet = [string]::Empty
    [string]$revison = [string]::Empty
    [string]$testTime = [string]::Empty
    [string]$firstFailed = [string]::Empty
    [string]$defectValue = [string]::Empty


    DataReader_HTML([string]$content) {
        if ([string]::IsNullOrEmpty($content) -eq $false) {
            $this.doc = $this.GettingDocument($content)
            $this.Update()


            ### clear invalid data
            if (($this.station -eq "ATS1" -or $this.station -eq "ATS2") -eq $false) {
                $this.revison = [string]::Empty
            }
        }
    }


    ##### private function
    hidden [object]GettingDocument($content) {
        $webBrowser = New-Object System.Windows.Forms.WebBrowser
        $webBrowser.ScriptErrorsSuppressed = 1
        $webBrowser.DocumentText = $content
        $webBrowser.Document.OpenNew(1)
        $webBrowser.Document.Write($content)
        $webBrowser.Refresh()
        return $webBrowser.Document
    } ### get document from html data

    hidden Update() {
        foreach($element in $this.doc.All) {
            if ([string]::IsNullOrEmpty($this.configName)) {
                $this.configName = $this.GetHTML_DATA_CELLNOCOLOR($element, "Configuration:")
                if (![string]::IsNullOrEmpty($this.configName)) {
                    $this.configName = $this.configName.Split("\\")[$this.configName.Split("\\").Length - 1]
                }

            }
            if ([string]::IsNullOrEmpty($this.pallet)) {
                $this.pallet = $this.GetHTML_DATA_CELLNOCOLOR($element, "Pallet_No:")
            }
            if ([string]::IsNullOrEmpty($this.tester)) {
                $this.tester = $this.GetHTML_DATA_CELLNOCOLOR($element, "Windows Computer Name:")
            }
            if ([string]::IsNullOrEmpty($this.station)) {
                $this.station = $this.GetHTML_DATA_CELLNOCOLOR($element, "Configuration:")
                if ($this.station.Contains("vFlash1")) {
                    $this.station = "Vflash1"
                }
            }
            if ([string]::IsNullOrEmpty($this.firstFailed)) {
                if (($element.GetAttribute("className") -eq "NegativeResultCell" -or $element.GetAttribute("className") -eq "InconclusiveResultCell") -and $element.TagName -eq "TD" -and $element.Children.Count -eq 1) {
                    foreach($ch in $element.Parent.Children) {
                        if ($ch.GetAttribute("className") -eq "DefaultCell" -and $ch.TagName -eq "TD" -and $ch.Children.Count -eq 1) {
                            $this.firstFailed = $ch.FirstChild.InnerText
                            foreach($chh in $ch.Parent.Children) {
                                if (($chh.GetAttribute("className") -eq "NegativeResultCell" -or $chh.GetAttribute("className") -eq "InconclusiveResultCell") -and $chh.TagName -eq "TD" -and $chh.Children.Count -eq 1) {
                                    $this.valueID = $chh.FirstChild.GetAttribute("href").Split("#")[$chh.FirstChild.GetAttribute("href").Split("#").Length - 1]
                                }
                            }
                            
                        }
                    }
                }
            }
            if (![string]::IsNullOrEmpty($this.valueID)) {
                if ($element.Id -eq $this.valueID) {
                    $idCount = 0
                    foreach($ch in $element.Parent.All) {
                        if (($ch.GetAttribute("className") -eq "NegativeResultCell" -or $ch.GetAttribute("className") -eq "InconclusiveResultCell") -and $ch.TagName -eq "TD") {
                            $getValue = [string]::Empty
                            $getCount = 0
                            foreach($chh in $ch.Parent.Children) {
                                if ($chh.GetAttribute("className") -eq "DefaultCell") {
                                    if ($getCount -eq 0) {
                                        $getValue = $chh.InnerText 
                                    }
                                    else {
                                        $getValue = [string]::Format('{0} {1}', $getValue, $chh.InnerText)
                                    }
                                    $getCount++
                                }
                            }
                            if ($idCount -eq 0) {
                                $this.defectValue = $getValue
                            }
                            else {
                                $this.defectValue = [string]::Format('{0}//{1}', $this.defectValue, $getValue)
                            }
                            $idCount++
                        }
                    }
                }
            }
        }
    } ### update data

    hidden [string]GetHTML_DATA_CELLNOCOLOR($elment, $containsWord) {
        $result = [string]::Empty
        if (![string]::IsNullOrEmpty($elment.InnerText)) {
            if ($elment.InnerText.Contains($containsWord) -and $elment.TagName -eq "TD" -and $elment.GetAttribute("className") -eq "CellNoColor") {
                $result = $elment.Parent.Children[$elment.Parent.Children.Count - 1].InnerText
            }
        }
        return $result
    } ### get data by node
}

##### data reader txt
class DataReader_TXT {
    ### private data
    hidden [Int32]$defectCount = 0
    hidden [Int32]$dataCount = 0
    hidden [string[]]$datas


    ### public data
    [string]$configName = [string]::Empty
    [string]$ext = [string]::Empty
    [string]$model = [string]::Empty
    [string]$station = [string]::Empty
    [string]$tester = [string]::Empty
    [string]$pallet = [string]::Empty
    [string]$revison = [string]::Empty
    [string]$testTime = [string]::Empty
    [string]$firstFailed = [string]::Empty
    [string]$defectValue = [string]::Empty
    #

    DataReader_TXT([string]$content) {
        if ([string]::IsNullOrEmpty($content) -eq $false) {
            $this.GettingDocument($content)
            
            $this.Update()
        }
    }


    ##### private function
    hidden GettingDocument($content) {
        $this.datas = $content.Split([System.Environment]::NewLine)
    } ### getting document
    hidden Update() {
        foreach($data in $this.datas) {
            if ([string]::IsNullOrEmpty($data.Trim()) -eq $false) {
                if ($data.Contains("Test Program Name :") -and [string]::IsNullOrEmpty($this.configName)) {
                    $data = $data.Trim()
                    foreach($d in $data.Split(" ")){
                        if ($d.Contains("EAP")) {
                            $this.configName = $d
                        }
                    }
                }
                if ($data.Contains("Test Station Name=Gen4") -and [string]::IsNullOrEmpty($this.station)) {
                    $data = $data.Trim()
                    $this.station = $data.Split(" ")[$data.Split(" ").Length - 1]
                }
                if ($data.Contains("Pallet No=") -and [string]::IsNullOrEmpty($this.pallet)) {
                    $data = $data.Trim()
                    $this.pallet = $data.Split("=")[$data.Split("=").Length - 1]
                }
                if ($data.Contains(" Seq    Mode     SpecHigh         SpecLow       Output        Measure       Result") -and [string]::IsNullOrEmpty($this.defectValue) -and $this.defectCount -eq 0) {
                    for($i = $this.dataCount; $i -lt ($this.dataCount + 25); $i++) {
                        if ([string]::IsNullOrEmpty($this.defectValue)) {
                            $this.defectValue = $this.datas[$i]
                        }
                        else {
                            $this.defectValue = [string]::Format('{0}//{1}', $this.defectValue, $this.datas[$i])
                        }
                        
                        $getData = $this.datas[$i].Trim()
                        if ($getData.Split(" ")[$getData.Split(" ").Length - 1] -eq "0" -and [string]::IsNullOrEmpty($this.firstFailed) -eq $true) {
                            if ($getData.Split(" ")[$getData.Split(" ").Length - 2] -eq "99999.00000") {
                                $this.firstFailed = [string]::Format("Step {0} {2}", $getData.Split(" ")[0], "OPEN")
                            }
                            else {
                                $this.firstFailed = [string]::Format("Step {0} FAILED", $getData.Split(" ")[0])
                            }
                        }
                        if ($getData.Split(" ")[0] -eq "8") {
                            $i = $this.datas.Count
                        }
                    }
                }
            }
            $this.dataCount++
        }
    } ### update data
}

[Program]$program = New-Object Program