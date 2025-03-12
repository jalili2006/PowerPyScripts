# Längenprüfung
# Skript zum Ermitteln der Längen von Dateinamen mit Pfaden
# Das Skrirpt ist auf das zu durchsuchende Startverzeichnis zu kopieren und kann dann von dort gestartet werden.
# Es werden rekursiv alle Ordner und Dateien ab aktuellem Pfad durchsucht und in der Logdatei mit Namen der Maschine gespeichert.
# Alle Dateien, deren Zeichenpfad zu lang ist, werden auf dem Bildschirm ausgegeben.

# Version: 1.5

$GetFiles = @"
using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

public class GetFiles
{
	internal static IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);
	internal static int FILE_ATTRIBUTE_DIRECTORY = 0x00000010;
	internal const int MAX_PATH = 260;

	[StructLayout(LayoutKind.Sequential)]
	internal struct FILETIME
	{
		internal uint dwLowDateTime;
		internal uint dwHighDateTime;
	};

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	internal struct WIN32_FIND_DATA
	{
		internal FileAttributes dwFileAttributes;
		internal FILETIME ftCreationTime;
		internal FILETIME ftLastAccessTime;
		internal FILETIME ftLastWriteTime;
		internal int nFileSizeHigh;
		internal int nFileSizeLow;
		internal int dwReserved0;
		internal int dwReserved1;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH)]
		internal string cFileName;
		// not using this
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
		internal string cAlternate;
	}

	[Flags]
	public enum EFileAccess : uint
	{
		GenericRead = 0x80000000,
		GenericWrite = 0x40000000,
		GenericExecute = 0x20000000,
		GenericAll = 0x10000000,
	}

	[Flags]
	public enum EFileShare : uint
	{
		None = 0x00000000,
		Read = 0x00000001,
		Write = 0x00000002,
		Delete = 0x00000004,
	}

	public enum ECreationDisposition : uint
	{
		New = 1,
		CreateAlways = 2,
		OpenExisting = 3,
		OpenAlways = 4,
		TruncateExisting = 5,
	}

	[Flags]
	public enum EFileAttributes : uint
	{
		Readonly = 0x00000001,
		Hidden = 0x00000002,
		System = 0x00000004,
		Directory = 0x00000010,
		Archive = 0x00000020,
		Device = 0x00000040,
		Normal = 0x00000080,
		Temporary = 0x00000100,
		SparseFile = 0x00000200,
		ReparsePoint = 0x00000400,
		Compressed = 0x00000800,
		Offline = 0x00001000,
		NotContentIndexed = 0x00002000,
		Encrypted = 0x00004000,
		Write_Through = 0x80000000,
		Overlapped = 0x40000000,
		NoBuffering = 0x20000000,
		RandomAccess = 0x10000000,
		SequentialScan = 0x08000000,
		DeleteOnClose = 0x04000000,
		BackupSemantics = 0x02000000,
		PosixSemantics = 0x01000000,
		OpenReparsePoint = 0x00200000,
		OpenNoRecall = 0x00100000,
		FirstPipeInstance = 0x00080000
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct SECURITY_ATTRIBUTES
	{
		public int nLength;
		public IntPtr lpSecurityDescriptor;
		public int bInheritHandle;
	}

	[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
	internal static extern IntPtr FindFirstFile(string lpFileName, out
									WIN32_FIND_DATA lpFindFileData);

	[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
	internal static extern bool FindNextFile(IntPtr hFindFile, out
									WIN32_FIND_DATA lpFindFileData);

	[DllImport("kernel32.dll", SetLastError = true)]
	[return: MarshalAs(UnmanagedType.Bool)]
	internal static extern bool FindClose(IntPtr hFindFile);

	// Assume dirName passed in is already prefixed with \\?\
	public static List<string> FindFilesAndDirs(string dirName)
	{
		List<string> results = new List<string>();
		WIN32_FIND_DATA findData;
		IntPtr findHandle = FindFirstFile(dirName + @"\*", out findData);
		if (findHandle != INVALID_HANDLE_VALUE)
		{
			bool found;
			do
			{
				string currentFileName = findData.cFileName;
				// if this is a directory, find its contents
				if (((int)findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
				{
					if (currentFileName != "." && currentFileName != "..")
                {
						List<string> childResults = FindFilesAndDirs(Path.Combine(dirName, currentFileName));
						// add children and self to results
						results.AddRange(childResults);
						results.Add(Path.Combine(dirName, currentFileName));
					}
				}
				// it’s a file; add it to the results
				else
				{
					results.Add(Path.Combine(dirName, currentFileName));
				}
				// find next
				found = FindNextFile(findHandle, out findData);
			}
			while (found);
		}
		// close the find handle
		FindClose(findHandle);
		return results;
	}
}
"@
Add-Type $GetFiles
Add-Type -AssemblyName System.Windows.Forms
## Funktions Deklaration ##

Function Start-Filesearch {
    param(
        [Parameter(Mandatory=$true)]
        [string]$pfad = (Get-Location),

        [Parameter(Mandatory=$true)]
        [string]$laenge
        )

    $laenge = $laenge - 4
    $files = [GetFiles]::FindFilesAndDirs("\\?\$pfad")

   $i = 0
   $a = 0
    forEach ($file in $files) {
        If ($file.Length -lt $laenge){

        #$count01 = ($files).count
          
        
     $i=$i+1
     #$link01="<a href=$($src) target=_explorer.exe> Link Text </a>"
       
 #$global:html += "<tr> <td>$($link01)</td> <td> $($file.TrimStart('\\?\')) </td> <td>$($file.Length)</td></tr>"


 $global:html += "<tr> <td>$($i)</td> <td> $($file.TrimStart('\\?\')) </td><td>$($file.Length)</td></tr>"


 #$global:htmlerroronly += "<tr><td><FONT COLOR=`"#FF0000`">$trim</FONT></td><td><FONT COLOR=`"#FF0000`">$trim</FONT></td><td><FONT COLOR=`"#FF0000`">$($file.Length-$maxlaenge)</FONT></td></tr>"
                      
        } 


        else {
        
        $a = $a+1

        $path_org = $file.TrimStart("\\?\")
           
        $path_ch = Split-Path -Path $path_org

        $link01="<a href=$($path_ch) target=_explorer.exe> Link Text </a>"

        

            $trim = $file.TrimStart('\\?\')
          $global:html += "<tr><td><FONT COLOR=`"#FF0000`">$trim</FONT></td><td><FONT COLOR=`"#FF0000`">$($file.Length)</FONT></td></tr>"
            
            #$global:html += $global:htmlerroronly.Replace()
            
            
            #$global:htmlerroronly += "<tr><td>$($trim)</td><td><FONT COLOR=`"#FF0000`">$($trim)</FONT></td><td>$($link01)</td><td><FONT COLOR=`"#FF0000`">$($file.Length-$maxlaenge)</FONT></td></tr>"
        #$global:htmlerroronly += "<tr><td>$($i)</td><td> $($trim) </td><td>$($link01)</td><td>$($file.Length)</td></tr>"
        
      $file01 = $file.Length
      $max01 = $maxlaenge

      
if($file01 -le $max01 ){

#$global:htmlerroronly += "<tr><td><FONT COLOR=`"#FF0000`"> $($a) </FONT></td><td><FONT COLOR=`"#FF0000`">$trim</FONT></td><td>$($link01)</td><td><FONT COLOR=`"#FF0000`">$($file.Length)</FONT></td></tr>"

 
#$global:html += "<tr> <td>$($i)</td> <td> $($file.TrimStart('\\?\')) </td><td>$($file.Length-$maxlaenge)</td></tr>"


} 

Else{

$global:htmlerroronly += "<tr><td><FONT COLOR=`"#FF0000`"> $($a) </FONT></td><td><FONT COLOR=`"#FF0000`">$trim</FONT></td><td>$($link01)</td><td><FONT COLOR=`"#FF0000`">$($file.Length-$maxlaenge)</FONT></td></tr>"


}


        }
    }
}

## Variablen Deklaration ##

$htmltablebottom = @"
</tbody></table>
</body>
</html>
"@

$htmlhead = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
    <title>Filesearch HTML Report</title>
    <meta name="robots" content="noindex, nofollow">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <style>
        body, html{ height:100%; width:100%; margin:0px; padding:0px; }
        body{ font-family:Verdana, Geneva, Arial, Helvetica, sans-serif; font-size:12px; }
        th{ background:#99bfe6; }
        tr:nth-child(even){ background:#f5f5f5; }
        th,td{ text-align:left; border:1px solid #bbb; word-wrap:break-word;}
        table {border-collapse:collapse; table-layout:fixed; width:100%;}
        table td {border:solid 1px #fab; word-wrap:break-word;}
   </style>
</head>
<body>
"@

$htmltablehead = @"
<table id="data-table" border="1" cellpadding="8" cellspacing="0">
<thead>
<tr>
<th width='10%' nowrap>ID </th> <th width='80%' nowrap>Datei</th> <th width='10%' nowrap>Pfadlaenge</th>
</tr>
</thead>
<tbody>
"@

$htmltableheaderror = @"
<table id="error-table" border="1" cellpadding="8" cellspacing="0">
<thead>
<tr>
<th width='5%' nowrap>ID </th> <th width='80%' nowrap>Datei</th> <th width='10%' nowrap>Link </th><th width='10%' nowrap>kürzendeLänge</th>
</tr>
</thead>
<tbody>
"@


#<th nowrap>Datei</th><th width='15%' nowrap>Zu kürzende Länge</th>

$htmlbreak = @"
<br></br>
"@


# wie viele Zeichen muss ich von der zu prüfenden Länge abziehen, Freigabe im Dateisystem auf dem Server?
$vorgabepfad = '\\cc03wsv0061\egovschulen$'
# Die Maximallänge ergibt sich aus der Maximallänge im Dateisystem abzüglich des Pfades der Unterordner
$maxlaenge = 256 - $vorgabepfad.Length



# Logpfad
$date = get-date -Format 'dd_MM_yy_hh_mm'
$global:LogPath = "Laengenpruefung-$(gc env:computername)-$date.html"

## Main start ##

$global:html = $htmlhead
#$global:html += "<h1>Report über die Dateipfadlänge vom [$([DateTime]::Now)]</h1>"

$date2 = get-date -Format 'dd.MM.yyyy HH:mm'
$global:html += "<h1>Report über die Dateipfadlänge vom [$($date2)]</h1>"

$global:html += $htmltablehead


$msgBox = New-Object -ComObject Shell.Application
$folder = $msgBox.BrowseForFolder(0, "Bitte Ordner wählen", 512)

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($msgBox) > $null
$src = $folder.Self.Path

Start-Filesearch -pfad $src -laenge $maxlaenge






#$url = "$src"
#$charCount = ($url.ToCharArray() | Where-Object {$_ } | Measure-Object).Count

#if ($charCount -ge 226){

#$global:html += "<h1>Anzahl der Fehlerdateien: [$("$charcount")]</h1>"

#}

#Else  {

#$global:html += "<h1>Anzahl der Fehlerdateien: [$("$charcount")]</h1>"

#}





$colItems = (Get-ChildItem -Path $src | Measure-Object -property length -sum)
$size_p = $colItems.Sum
$size_mb = ($size_p / 1MB)
$size_gb = ($size_p / 1GB)


$final_s_m =[math]::Round($size_mb,2)
$final_s_g = [math]::Round($size_gb,2)

$global:html += "<h1>Gesamtgröße [$($final_s_m)] MegaByte + [$($final_s_g)] Gigabyte </h1>"


$num01 = (get-childitem -Path $src -recurse | where-object { $_.FullName }).Count
$num01_error = (get-childitem -Path $src -recurse | where-object { $_.FullName }).Count
#$global:html += "<h1>Number of files in folder: [$($num01)]</h1>"

#Write-Host $num01






$global:html += $htmltablebottom
$global:html += $htmlbreak
$global:html += '<h1>Zusammenfassung der zu langen Dateipfade</h1>'
$global:html += $htmltableheaderror
$global:html += $global:htmlerroronly 
$global:html += $htmltablebottom
$global:html | Out-File "$(Get-Location)\$global:LogPath"
#Invoke-Item "$(Get-Location)\$global:LogPath"
Invoke-Item "$(Get-Location)\$global:LogPath" 

