# Created By: Michael Rath und Sebastian Trebbin, smapOne AG
# Version 1.6
#
# Änderungen:
# * Erläuterung für die Konfiguration eingepflegt
# * Download von Assets eines Datensatzes
# * Erweiterung des HTTP Requests um API Limitierung (HTTP Status 429)
#
#
# Systemvoraussetzungen: 
# * das Skript kann nur unter PowerShell Version 5.1.x und höher korrekt betrieben werden
# * über den Befehl $PSVersionTable kann dies in einer PowerShell Console eingesehen werden
# * für Windows Server 2016 sowie Windows 10 Systeme sollte dies die vorinstallierte Version sein
# * für Windows Server 2012 sowie Windows 7/8 kann das Update über folgenden Link bezogen werden: https://docs.microsoft.com/de-de/powershell/wmf/5.1/install-configure
#
#
# KONFIGURATION: 
#
# das Config-Array regelt alle Einstellungen für den Zugriff auf die smapOne Apps, den Datenbezug etc.
# für jede smapOne App muss ein separater Eintrag in Array mit den jeweils gewünschten Einstellungen vorgenommen werden
# Parameter des Arrays: 
#    smapId - die Id der smapOne App, welche u.a. aus der Browser URL für die App entnommen werden kann, Datentyp: GUID
#    exportJSON - soll der Datensatz als JSON exportiert werden, Datentyp: Bool ($TRUE, $FALSE)
#    exportJSONPath - Dateipfad unter dem die JSON-Datei abgelegt werden soll, UNC-Pfade sind ebenfalls möglich, Datentyp: String
#    exportJSONName - Dateiname unter dem die JSON-Datei abgelegt werden soll, kann dynamisch durch Inhalte des Datensatzes ergänzt werden, siehe separate Beschreibung
#    exportPDF - soll der Datensatz als PDF exportiert werden, Datentyp: Bool ($TRUE, $FALSE)
#    exportPDFPath - Dateipfad unter dem die PDF-Datei abgelegt werden soll, UNC-Pfade sind ebenfalls möglich, Datentyp: String
#    exportPDFNameUsePlattform - der Dateiname des PDF-Berichtes wird direkt aus den Einstellungen der Plattform entnommen, Datentyp: Bool ($TRUE, $FALSE)
#    exportPDFName - Dateiname unter dem die PDF-Datei abgelegt werden soll, kann dynamisch durch Inhalte des Datensatzes ergänzt werden, siehe separate Beschreibung
#    exportAssets - die dem Datensatz enthaltenen Fotos / Skizzen etc. werden separat heruntergeladen
#    exportAssetsPath - Dateipfad unter dem die separaten Assets abgelegt werden sollen, UNC-Pfade sind ebenfalls möglich, Datentyp: String
#    deleteAfterExport - Datensätze nach dem Export löschen
#    token - Login Token des Creator Accounts unter welchem die App abgelegt ist, Datentyp: GUID
#    debug - abgerufene Datensätze werden nicht als exportiert markiert
#
# Erweiterungen für Dateinamen
#    ein Dateiname kann eine einfacher String sein, Beispiel: exportJSONName='JSON Bericht Wartung.json'
#    ein Dateiname kann direkt aus der Benennung des Berichtes der Plattform resultieren
#    
#-------------------------------------------------------------------------------

$Logfile = "$PSScriptRoot\log.txt"
$Proxy = ""

$configArray = @(
    # Beispiel: 
    [pscustomobject]@{
        smapId = "00000000-dead-beef-0000-000000000123";

        exportJSON = $true;
        exportJSONPath = "$PSScriptRoot\export";
        exportJSONName = 'JSON_$($R.id).json'
        
        exportPDF = $true;
        exportPDFPath = '$PSScriptRoot\export\';
        exportPDFNameUsePlattform = $true;
        exportPDFName = '';
        
        exportAssets = $true;
        exportAssetsPath = '$PSScriptRoot\$($R.id)';

        deleteAfterExport = $false;
        
        token = "BeispielToken";
        debug = $false;
    }
    #,[pscustomobject]@{
    #    smapId = "8c76e6c6-3186-4dfd-9eab-f5e4e5975f77";
    #
    #    exportJSON = $false;
    #    exportJSONPath = "$env:USERPROFILE\Desktop\export";
    #    exportJSONName = 'JSON Test $($R.id).json'
    #    
    #    exportPDF = $false;
    #    exportPDFPath = '$env:USERPROFILE\Desktop\export\';
    #    exportPDFNameUsePlattform = $true;
    #    exportPDFName = '';
    #
    #    deleteAfterExport = $false;
    #    
    #    exportAssets = $true;
    #    exportAssetsPath = '$env:USERPROFILE\Desktop\export\$($R.id)';
    #    
    #    token = "BEISPIELTOKEN";
    #    debug = $false;
    #}
) 

function Log($Text) {

    $Text = "$(Get-Date -Format g) $Text";
    Write-Host $Text;

    Add-content $Logfile -value $Text
}

function GetFileName($fileName) {
    $fileName = $fileName.Replace("inline; filename=", "")
    if ($fileName.IndexOf("=?utf-8?B?") -ge 0) {                    
        $fileName = $fileName.Replace("=?utf-8?B?", "")
        $fileName = $fileName.Replace("?=", "")
        $fileName = $fileName.Substring(1, $fileName.length - 2) #da führende und Ende " entfernen
        $fileName = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($fileName))                    
    } else {
        if ($fileName.StartsWith("""")) {
            $fileName = $fileName.Substring(1, $fileName.length - 2)
        }
    } 
    return $fileName 
}


function doWebRequest ($uriString) {
    
    try {
        if (($Proxy -ne "") -and ($Proxy -ne $null)) {
            $Response = Invoke-WebRequest -Proxy $Proxy $uriString         
        } else {
            $Response = Invoke-WebRequest $uriString        
        }
        return $Response;
    } catch {
        if ($_.Exception.Response.StatusCode -eq 429) {
            $waitSecs = $_.Exception.Response.Headers["Retry-After"];            
            Log("HTTP Status 429 - wait $waitSecs Seconds");
            Start-Sleep -s $waitSecs;
            return doWebRequest -uriString $uriString;
        } else {
            Log("Fehler: $($_.Exception.Message) - $($_.Exception.Response.StatusCode)")
            return "";
        }

    }
}


function doWebRequestDelete ($uriString) {
    
    try {
        $Response = Invoke-WebRequest -Method Delete $uriString        
        return $Response;
    } catch {
        if ($_.Exception.Response.StatusCode -eq 429) {
            $waitSecs = $_.Exception.Response.Headers["Retry-After"];            
            Log("HTTP Status 429 - wait $waitSecs Seconds");
            Start-Sleep -s $waitSecs;
            return doWebRequest -uriString $uriString;
        } else {
            Log("Fehler: $($_.Exception.Message) - $($_.Exception.Response.StatusCode)")
            return "";
        }
    }
}

function doWebRequestDownloadFile($uriString, $path, $fileName) {
        
    if ($null -eq $fileName) {
        $fileMeta = doWebRequest -uriString $uriString;
        $fileName = GetFileName($fileMeta.headers["Content-Disposition"])
    }
    
    Log("Write File: $($path)\$($fileName)")

    try {                  
        if (($Proxy -ne "") -and ($Proxy -ne $null)) {
            $Response = Invoke-WebRequest $uriString -Proxy $Proxy -OutFile "$($path)\$($fileName)"
        } else {
            $Response = Invoke-WebRequest $uriString -OutFile "$($path)\$($fileName)"
        }
    } catch {
        if ($_.Exception.Response.StatusCode -eq 429) {
            $waitSecs = $_.Exception.Response.Headers["Retry-After"];
            Log("wait $waitSecs Seconds");
            Start-Sleep -s $waitSecs;
            doWebRequestDownloadFile -uriString $uriString -path $path -fileName $fileName
        } else {
            Log("Fehler: $($_.Exception.Message) - $($_.Exception.Response.StatusCode)")
            return "";
        } 
    } 
}


function GetRec($EC, $R) {
    
    Log("Processing Record: $($R.id)")

    if ($EC.exportJSON) {
        $exportJSONPath = $ExecutionContext.InvokeCommand.ExpandString($EC.exportJSONPath);
        if (!(Test-Path $exportJSONPath)) {
            Log("Create Folder: $($exportJSONPath)")
            New-Item -ItemType Directory -Force -Path $exportJSONPath | Out-Null
        }
        Log("Export JSON")
        $httpString = "https://platform.smapone.com/backend/intern/Smaps/$($R.smapId)/Versions/$($R.version)/Data/$($R.id).json?markAsExported=false&useDefault=false&accesstoken=$($EC.token)"
        $jsonName = $ExecutionContext.InvokeCommand.ExpandString($EC.exportJSONName) -replace "`n", "" -replace "`r", ""; 
        doWebRequestDownloadFile -uriString $httpString -path $EC.exportJSONPath -fileName $jsonName
        
    }

    if ($EC.exportPDF) {
        $exportPDFPath = $ExecutionContext.InvokeCommand.ExpandString($EC.exportPDFPath);
        if (!(Test-Path $exportPDFPath)) {
            Log("Create Folder: $($exportPDFPath)")
            New-Item -ItemType Directory -Force -Path $exportPDFPath | Out-Null
        }

        Log("Export PDF")
        $httpString = "https://platform.smapone.com/backend/intern/Smaps/$($R.smapId)/Versions/$($R.version)/Data/$($R.id).Pdf?markAsExported=false&useDefault=false&accesstoken=$($EC.token)"
            
        $pdfName = $null;
        if ($EC.exportPDFNameUsePlattform -eq $false) {
            $pdfName = $ExecutionContext.InvokeCommand.ExpandString($EC.exportPDFName) -replace "`n", "" -replace "`r", "";            
        } 

        doWebRequestDownloadFile -uriString $httpString -path $exportPDFPath -fileName $pdfName
    }

    if ($EC.exportAssets) {
        $exportAssetsPath = $ExecutionContext.InvokeCommand.ExpandString($EC.exportAssetsPath);
        if (!(Test-Path $exportAssetsPath)) {
            Log("Create Folder: $($exportAssetsPath)")
            New-Item -ItemType Directory -Force -Path $exportAssetsPath | Out-Null
        }

        Log("Export Assets")
        $httpString = "https://platform.smapone.com/backend/intern/Smaps/$($R.smapId)/Versions/$($R.version)/Data/$($R.id)/Files?accesstoken=$($EC.token)"
        $httpResultList = doWebRequest -uriString $httpString;    
        $jsonList = $httpResultList.Content | ConvertFrom-Json
                
        foreach ($jsonRec in $jsonList) {
            $httpString = "https://platform.smapone.com/backend/intern/Smaps/$($R.smapId)/Versions/$($R.version)/Data/$($R.id)/Files/$($jsonRec.fileId)?accesstoken=$($EC.token)"
            $assetName = $jsonRec.fileName;
            if (!($assetName.StartsWith("Signature_"))) {
                # Log("Write File: $($exportAssetsPath)\$($assetName)")
                doWebRequestDownloadFile -uriString $httpString -path $exportAssetsPath -fileName $assetName                
            }
        }
    }

    if ($EC.deleteAfterExport -and -not $EC.debug) {
        Log("Delete $($R.id)")
        $httpString = "https://platform.smapone.com/backend/intern/Smaps/$($R.smapId)/Versions/$($R.version)/Data/$($R.id)?accesstoken=$($EC.token)"
        doWebRequestDelete -uriString $httpString;
    }
}


function GetRecList($ElementConfig) {

    Log("Start loading Recordlist with Config")
    Log("SmapId: $($ElementConfig.smapId)")
    
    if ($ElementConfig.debug) {
        $httpString = "https://platform.smapone.com/backend/intern/Smaps/$($ElementConfig.smapId)/Data?markAsExported=false&format=Json&accessToken=$($ElementConfig.token)"
    } else {
        $httpString = "https://platform.smapone.com/backend/intern/Smaps/$($ElementConfig.smapId)/Data?markAsExported=true&format=Json&state=New&accessToken=$($ElementConfig.token)"
    }
    
    $httpResultList = doWebRequest -uriString $httpString;    
    $jsonList = $httpResultList.Content | ConvertFrom-Json    
    Log("Element count: $($jsonList.Count)")

    foreach ($jsonRec in $jsonList) {
        GetRec -EC $ElementConfig -R $jsonRec
        Log("++++++++++")
    }
    Log("##########")
}


function StartRoutine($Config) {
    Log("----------------")
    Log("Start Script")

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

    foreach ($element in $Config) {
        GetRecList($element)
    }
}

StartRoutine($configArray)
