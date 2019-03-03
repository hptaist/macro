function Get-RandomCharacters($length, $characters) { 
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
    $private:ofs="" 
    return [String]$characters[$random]
}
function Invoke-SplitVBA{

    Param(
    [String]$InputFile,
    [String]$OutputFile
    )

    if (!(Test-Path -LiteralPath "$InputFile")){
        throw $InputFile + "does not exist"
    }

    if ((Test-Path -LiteralPath "$OutputFile")){
        $Continue = Read-Host -Prompt "Output file already exists, overwrite? [Y/n]"

        switch ($Continue) {
        "" {continue}
        "y" {continue}
        "Y" {continue}
        "n" {throw "Aborting"; exit}
        "N" {throw "Aborting"; exit}
        default {throw "Aborting"; exit}
        }

        Remove-Item $OutputFile
    }
    certutil.exe -encode $InputFile "tmp.txt"
    $VBA = Get-Content -Path "tmp.txt" -Raw
    $VBA = $VBA.Replace("`r`n","")
    $VBA = $VBA.Replace("-----BEGIN CERTIFICATE-----","")
    $VBA = $VBA.Replace("-----END CERTIFICATE-----","")
    $i = 0
    $scname = Get-RandomCharacters -length 10 -characters 'abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ'
    $sc = $scname + " = " + $scname

    while ($i -le ($VBA.Length-900)){
        $temp = (Get-RandomCharacters -length 10 -characters 'abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ')
        "Dim " + "$temp" + " As String" | Out-File -FilePath $OutputFile -Append
        $temp + " = " + '"' + $VBA.Substring($i,900) + '"' | Out-File -FilePath $OutputFile -Append
        $sc = $sc + " & " + $temp
        $i += 900
    }
    $temp = (Get-RandomCharacters -length 10 -characters 'abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ') 
    "Dim " + "$temp" + " As String" | Out-File -FilePath $OutputFile -Append
    $temp + " = " + '"' + $VBA.Substring($i) + '"' | Out-File -FilePath $OutputFile -Append
    $sc = $sc + " & " + $temp
    $sc | Out-File -FilePath $OutputFile -Append
    Write-Host "Done"
    Remove-Item "tmp.txt"

}