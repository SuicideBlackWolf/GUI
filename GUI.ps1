$ErrorActionPreference= 'silentlycontinue'
Add-Type -Path "C:\Eskulap\orant\BIN\OdtPrinting\Oracle.ManagedDataAccess.dll"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '731,445'
$Form.text                       = "Przygotował Michał Zbyl. Ver. 2.4.2"
$form.StartPosition              = "centerscreen"
$Form.TopMost                    = $false

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Hasło"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(25,38)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$Button5                         = New-Object system.Windows.Forms.Button
$Button5.text                    = "Odblokuj"
$Button5.width                   = 80
$Button5.height                  = 30
$Button5.location                = New-Object System.Drawing.Point(25,68)
$Button5.Font                    = 'Microsoft Sans Serif,10'

$Button6                         = New-Object system.Windows.Forms.Button
$Button6.text                    = "Certyfikat"
$Button6.width                   = 80
$Button6.height                  = 30
$Button6.location                = New-Object System.Drawing.Point(106,68)
$Button6.Font                    = 'Microsoft Sans Serif,10'

$Button13                         = New-Object system.Windows.Forms.Button
$Button13.text                    = "Cer-Usuń"
$Button13.width                   = 80
$Button13.height                  = 30
$Button13.location                = New-Object System.Drawing.Point(106,98)
$Button13.Font                    = 'Microsoft Sans Serif,10'

$Button9                         = New-Object system.Windows.Forms.Button
$Button9.text                    = "Szukaj loginu po nazwisku"
$Button9.width                   = 100
$Button9.height                  = 30
$Button9.location                = New-Object System.Drawing.Point(186,68)
$Button9.Font                    = 'Microsoft Sans Serif,8'

$Button23                         = New-Object system.Windows.Forms.Button
$Button23.text                    = "Szukaj loginu po p0"
$Button23.width                   = 80
$Button23.height                  = 30
$Button23.location                = New-Object System.Drawing.Point(106,210)
$Button23.Font                    = 'Microsoft Sans Serif,8'

$Button27                         = New-Object system.Windows.Forms.Button
$Button27.text                    = "SID po p0"
$Button27.width                   = 80
$Button27.height                  = 25
$Button27.location                = New-Object System.Drawing.Point(106,240)
$Button27.Font                    = 'Microsoft Sans Serif,8'

$Button15                         = New-Object system.Windows.Forms.Button
$Button15.text                    = "Status Pacjenta po nazwisku"
$Button15.width                   = 100
$Button15.height                  = 30
$Button15.location                = New-Object System.Drawing.Point(286,98)
$Button15.Font                    = 'Microsoft Sans Serif,8'

$Button26                         = New-Object system.Windows.Forms.Button
$Button26.text                    = "Log Użyszkodnika po loginie"
$Button26.width                   = 100
$Button26.height                  = 30
$Button26.location                = New-Object System.Drawing.Point(186,98)
$Button26.Font                    = 'Microsoft Sans Serif,8'

$Button7                         = New-Object system.Windows.Forms.Button
$Button7.text                    = "Konsultacje"
$Button7.width                   = 70
$Button7.height                  = 25
$Button7.location                = New-Object System.Drawing.Point(288,42)
$Button7.Font                    = 'Microsoft Sans Serif,8'

$Button8                         = New-Object system.Windows.Forms.Button
$Button8.text                    = "Blokada Wypisy"
$Button8.width                   = 80
$Button8.height                  = 30
$Button8.location                = New-Object System.Drawing.Point(386,98)
$Button8.Font                    = 'Microsoft Sans Serif,8'

$Button25                         = New-Object system.Windows.Forms.Button
$Button25.text                    = "Pacjent Info"
$Button25.width                   = 80
$Button25.height                  = 30
$Button25.location                = New-Object System.Drawing.Point(286,68)
$Button25.Font                    = 'Microsoft Sans Serif,8'

$Button28                         = New-Object system.Windows.Forms.Button
$Button28.text                    = "ID Operacji"
$Button28.width                   = 80
$Button28.height                  = 30
$Button28.location                = New-Object System.Drawing.Point(366,68)
$Button28.Font                    = 'Microsoft Sans Serif,8'

$Button24                         = New-Object system.Windows.Forms.Button
$Button24.text                    = "Ponowne skierowanie"
$Button24.width                   = 80
$Button24.height                  = 30
$Button24.location                = New-Object System.Drawing.Point(480,38)
$Button24.Font                    = 'Microsoft Sans Serif,8'

$TextBox4                        = New-Object system.Windows.Forms.TextBox
$TextBox4.multiline              = $false
$TextBox4.width                  = 60
$TextBox4.height                 = 20
$TextBox4.location               = New-Object System.Drawing.Point(368,44)
$TextBox4.Font                   = 'Microsoft Sans Serif,10'

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 150
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(110,44)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

$TextBox3                        = New-Object system.Windows.Forms.TextBox
$TextBox3.multiline              = $false
$TextBox3.width                  = 150
$TextBox3.height                 = 20
$TextBox3.location               = New-Object System.Drawing.Point(110,156)
$TextBox3.Font                   = 'Microsoft Sans Serif,10'

$Button22                         = New-Object system.Windows.Forms.Button
$Button22.text                    = "Połącz"
$Button22.width                   = 60
$Button22.height                  = 25
$Button22.location                = New-Object System.Drawing.Point(260,156)
$Button22.Font                    = 'Microsoft Sans Serif,10'

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $True
$TextBox2.width                  = 700
$TextBox2.height                 = 150
$TextBox2.location               = New-Object System.Drawing.Point(15,266)
$TextBox2.Font                   = 'Microsoft Sans Serif,10'
$TextBox2.ScrollBars             = "Vertical"
$textBox2.ReadOnly               = $True
$textbox2.anchor                 = [System.Windows.Forms.AnchorStyles]::Top `
                                    -bor [System.Windows.Forms.AnchorStyles]::Left `
                                    -bor [System.Windows.Forms.AnchorStyles]::Right `
                                    -bor [System.Windows.Forms.AnchorStyles]::Bottom

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Eskulap"
$Label1.AutoSize                 = $true
$Label1.width                    = 200
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(10,10)
$Label1.Font                     = 'Microsoft Sans Serif,15,style=Bold,Underline'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "RTG"
$Label3.AutoSize                 = $true
$Label3.width                    = 200
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(466,10)
$Label3.Font                     = 'Microsoft Sans Serif,15,style=Bold,Underline'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Domena"
$Label2.AutoSize                 = $true
$Label2.width                    = 200
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(10,120)
$Label2.Font                     = 'Microsoft Sans Serif,15,style=Bold,Underline'

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Wyczyść"
$Button2.width                   = 71
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(15,233)
$Button2.Font                    = 'Microsoft Sans Serif,10'

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.text                    = "Hasło"
$Button3.width                   = 60
$Button3.height                  = 30
$Button3.location                = New-Object System.Drawing.Point(25,150)
$Button3.Font                    = 'Microsoft Sans Serif,10'

$Button4                         = New-Object system.Windows.Forms.Button
$Button4.text                    = "Odblokuj"
$Button4.width                   = 80
$Button4.height                  = 30
$Button4.location                = New-Object System.Drawing.Point(25,180)
$Button4.Font                    = 'Microsoft Sans Serif,10'

$Button10                         = New-Object system.Windows.Forms.Button
$Button10.text                    = "Szukaj loginu po nazwisku"
$Button10.width                   = 80
$Button10.height                  = 30
$Button10.location                = New-Object System.Drawing.Point(106,180)
$Button10.Font                    = 'Microsoft Sans Serif,8'

$Button11                         = New-Object system.Windows.Forms.Button
$Button11.text                    = "Adres IP"
$Button11.width                   = 80
$Button11.height                  = 30
$Button11.location                = New-Object System.Drawing.Point(186,180)
$Button11.Font                    = 'Microsoft Sans Serif,10'

$Button14                         = New-Object system.Windows.Forms.Button
$Button14.text                    = "Adres IP - Wszystkie"
$Button14.width                   = 80
$Button14.height                  = 30
$Button14.location                = New-Object System.Drawing.Point(186,210)
$Button14.Font                    = 'Microsoft Sans Serif,8'

$Button20                         = New-Object system.Windows.Forms.Button
$Button20.text                    = "ping"
$Button20.width                   = 80
$Button20.height                  = 25
$Button20.location                = New-Object System.Drawing.Point(186,240)
$Button20.Font                    = 'Microsoft Sans Serif,8'

$Button21                         = New-Object system.Windows.Forms.Button
$Button21.text                    = "MAC/IP"
$Button21.width                   = 80
$Button21.height                  = 30
$Button21.location                = New-Object System.Drawing.Point(266,180)
$Button21.Font                    = 'Microsoft Sans Serif,8'

$Button12                         = New-Object system.Windows.Forms.Button
$Button12.text                    = "Email"
$Button12.width                   = 80
$Button12.height                  = 30
$Button12.location                = New-Object System.Drawing.Point(266,210)
$Button12.Font                    = 'Microsoft Sans Serif,10'

$Button16                         = New-Object system.Windows.Forms.Button
$Button16.text                    = "Ostatni PC"
$Button16.width                   = 100
$Button16.height                  = 25
$Button16.location                = New-Object System.Drawing.Point(346,240)
$Button16.Font                    = 'Microsoft Sans Serif,8'

$Button18                         = New-Object system.Windows.Forms.Button
$Button18.text                    = "AD - PC Raport"
$Button18.width                   = 100
$Button18.height                  = 30
$Button18.location                = New-Object System.Drawing.Point(346,210)
$Button18.Font                    = 'Microsoft Sans Serif,8'

$Button19                         = New-Object system.Windows.Forms.Button
$Button19.text                    = "AD - Użytkownicy Raport"
$Button19.width                   = 100
$Button19.height                  = 30
$Button19.location                = New-Object System.Drawing.Point(346,180)
$Button19.Font                    = 'Microsoft Sans Serif,8'

$ListBox                          = New-Object system.Windows.Forms.ListBox
$ListBox.text                     = "listBox"
$ListBox.width                    = 210
$ListBox.height                   = 180
$ListBox.location                 = New-Object System.Drawing.Point(480,80)

$Form.controls.AddRange(@($Button1,$ListBox,$Label3,$TextBox1,$Button25,$Button28,$Button27,$Button24,$Button26,$TextBox4,$Button16,$Button23,$Button22,$Button21,$Button19,$Button20,$Button18,$TextBox2,$Label1,$Button2,$Button15,$Button3,$Button13,$Button14,$Label2,$TextBox3,$Button4,$Button5,$Button6,$Button7,$Button12,$Button8,$Button9,$Button10,$Button11))

$userr = $env:UserName
$p1 = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(""))
$p2 = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(""))
$p3 = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(""))
$p4 = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(""))
$p5 = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(""))

if (($userr -eq $p1) -or ($userr -eq $p2) -or ($userr -eq $p3) -or ($userr -eq $p4) -or ($userr -eq $p5)) {
    $zuo1 = "zuo"
} else {
    Import-Module ActiveDirectory
    $group = "informatycy"
    $members = Get-ADGroupMember -Identity $group -Recursive | Select-Object -ExpandProperty SamAccountName
}

If (($members -contains $userr) -or ($zuo1 -eq "zuo")) {
    $l = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("="))
    $p = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("=="))
    ## To connect by SID
    $ora_server = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(""))
    $ora_user = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("="))
    $ora_pass = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("=="))
    $ora_sid = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("=="))


#ListBox
$ListBox.Items.AddRange()
$ListBox.Add_DoubleClick({
    $a = $ListBox.SelectedItem
    if ($a -like "*-*") {
        $c = $a.split('-')[1]
        $b = $c.split(' ')[1]
        $TextBox3.Text = $b
    } else {
        $TextBox3.Text = $a
    }
})

# SID po P
$Button27.Add_Click({
    $TextBox2.AppendText("Czekamy Cierpliwie...`r`n")

    $SID = $TextBox3.Text

    $objUser = New-Object System.Security.Principal.NTAccount("SZPITAL.LOCAL", $SID)
    $SIDF = $objUser.Translate([System.Security.Principal.SecurityIdentifier])

    $TextBox2.AppendText("`r`n")
    $check_userAD = Get-ADUser -Identity $SID | Select-Object -ExpandProperty Name
    $TextBox2.AppendText("$check_userAD - $SID`r`n")
    $TextBox2.AppendText("$SIDF`r`n")
})

# ID Operacji
$Button28.Add_Click({
    $nazwai = $TextBox1.Text

    if ($nazwai) {
        $TextBox2.AppendText("Informacje o ID Operacji na podstawie ID Pacjenta`r`n")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Możesz zawęzić dodając do nazwiska imię. Np. Barański Michał`r`n")
        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")
        ## by SID
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
        $query = "select * from ri_pacjenci where upper(p_nazwisko||' '||p_imie) like upper('%'||'$nazwai'||'%')"
        
        $connection.open()
        
        $command=$connection.CreateCommand()
        $command.CommandText=$query
        $wynik = $command.ExecuteReader()
        
        $table = new-object System.Data.DataTable
        $table.Load($wynik)
        
        $connection.close()
            
        $spr = $table 
        $sprfin = $spr | Out-GridView -PassThru -Title "Przygotował Michał Zbyl" | select-Object -ExpandProperty 'P_PACJENT_ID'

        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("ID Pacjenta: $sprfin`r`n")

        if ($sprfin) {
            $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
            $query = "select * from od_lecz_pacjenta where lec_p_pacjent_id = '$sprfin'"
            
            $connection.open()
            
            $command=$connection.CreateCommand()
            $command.CommandText=$query
            $wynik = $command.ExecuteReader()
            
            $table = new-object System.Data.DataTable
            $table.Load($wynik)
            
            $connection.close()
                
            $spr = $table 
            $sprfinf = $spr | Out-GridView -PassThru -Title "Przygotował Michał Zbyl" | select-Object -ExpandProperty 'LEC_LECZENIE_ID'

            $TextBox2.AppendText("ID Operacji: $sprfinf`r`n")
        }
    } else {
        $TextBox2.AppendText("Informacje o Pacjencie z bazy`r`n")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Pole puste`r`n")
        $TextBox2.AppendText("Możesz zawęzić dodając do nazwiska imię. Np. Barański Michał`r`n")
    }
})

# AD - PC Raport
$Button18.Add_Click({
    $TextBox2.AppendText("Zawsze Aktualne`r`n")
    $TextBox2.AppendText("Czekamy Cierpliwie aż się pokaże...`r`n")

    $audit1 = Get-ADComputer -Filter "ObjectClass -eq 'Computer'" `
        -Properties Name, OperatingSystem, `
            OperatingSystemServicePack, PasswordLastSet, IPv4Address,`
            whenCreated, whenChanged, LastLogonTimestamp,  `
            DistinguishedName |
        Where-Object {$_.whenChanged -gt $((Get-Date).AddDays(-90))} |
        Select-Object Name, OperatingSystem, `
            OperatingSystemServicePack, PasswordLastSet, IPv4Address,`
            whenCreated, whenChanged, `
            @{name='LastLogonTimestampDT';`
            Expression={[datetime]::FromFileTimeUTC($_.LastLogonTimestamp)}}, `
            DistinguishedName | Sort-Object Name

    $audit1 | Out-GridView -Title "Przygotował Michał Zbyl"
    $TextBox2.AppendText("`r`n")
})
# Ponowne Skierowanie
$Button24.Add_Click({
    $nazwau = $TextBox1.Text

    if ($nazwau) {
        $TextBox2.AppendText("Ponowne przesłanie z Esku do opisu`r`n")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Możesz zawęzić dodając do nazwiska imię. Np. Barański Michał`r`n")
        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")
        ## by SID
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
        $query = "select * from rtg_skierowania, ri_pacjenci where p_pacjent_id = s_pl_p_lokalny_id and upper(p_nazwisko||' '||p_imie) like upper('%'||'$nazwau'||'%')"
        
        $connection.open()
        
        $command=$connection.CreateCommand()
        $command.CommandText=$query
        $wynik = $command.ExecuteReader()
        
        $table = new-object System.Data.DataTable
        $table.Load($wynik)
        
        $connection.close()
            
        $spr = $table 
        $sprwynn = $spr | Out-GridView -PassThru -Title "Przygotował Michał Zbyl" | select-Object -ExpandProperty 'S_SKIEROWANIE_ID'

        if ($sprwynn) {
            ## To connect by SID
            $query2 = "delete from hl7_flds where hfl_symbol = 'zlec_id' and hfl_value like 'RTG.'||$sprwynn||'.%'"
            $query3 = "update rtg_skierowania set  s_status = 'R' where s_skierowanie_id = $sprwynn"
            $connection.open()
    
            $command2=$connection.CreateCommand()
            $command2.CommandText=$query2
            $command2.ExecuteReader()
    
            $connection.close()
            $connection.open()
    
            $command3=$connection.CreateCommand()
            $command3.CommandText=$query3
            $command3.ExecuteReader()
    
            $connection.close()

            $TextBox2.AppendText("Przesłano ponownie skierowanie $sprwynn`r`n")
        }
    } else {
        $TextBox2.AppendText("Ponowne przesłanie z Esku do opisu`r`n")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Pole puste`r`n")
        $TextBox2.AppendText("Możesz zawęzić dodając do nazwiska imię. Np. Barański Michał`r`n")
    }
})
# Pacjent - Info
$Button25.Add_Click({
    $nazwai = $TextBox1.Text

    if ($nazwai) {
        $TextBox2.AppendText("Informacje o Pacjencie z bazy`r`n")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Możesz zawęzić dodając do nazwiska imię. Np. Barański Michał`r`n")
        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")
        ## by SID
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
        $query = "select * from ri_pacjenci where upper(p_nazwisko||' '||p_imie) like upper('%'||'$nazwai'||'%')"
        
        $connection.open()
        
        $command=$connection.CreateCommand()
        $command.CommandText=$query
        $wynik = $command.ExecuteReader()
        
        $table = new-object System.Data.DataTable
        $table.Load($wynik)
        
        $connection.close()
            
        $spr = $table 
        $sprfin = $spr | Select-Object * | Out-GridView -PassThru -Title "Przygotował Michał Zbyl"

        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("$sprfin`r`n")
    } else {
        $TextBox2.AppendText("Informacje o Pacjencie z bazy`r`n")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Pole puste`r`n")
        $TextBox2.AppendText("Możesz zawęzić dodając do nazwiska imię. Np. Barański Michał`r`n")
    }
})
# Pacjent - Log
$Button26.Add_Click({
    $nazwal = $TextBox1.Text

    if ($nazwal) {
        $TextBox2.AppendText("Informacje o Użytkowniku`r`n")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")
        ## by SID
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
        $query = "SELECT * FROM STATS`$USER_LOG WHERE USER_ID = '$nazwal'"
        
        $connection.open()
        
        $command=$connection.CreateCommand()
        $command.CommandText=$query
        $wynik = $command.ExecuteReader()
        
        $table = new-object System.Data.DataTable
        $table.Load($wynik)
        
        $connection.close()
            
        $spr = $table 
        $sprfin = $spr | Out-GridView -PassThru -Title "Przygotował Michał Zbyl"

        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("$sprfin`r`n")
    } else {
        $TextBox2.AppendText("Informacje o Użytkowniku`r`n")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Pole puste`r`n")
    }
})
# AD - Użytkownicy Raport
$Button19.Add_Click({
    $TextBox2.AppendText("Zawsze Aktualne`r`n")
    $TextBox2.AppendText("Czekamy Cierpliwie aż się pokaże...`r`n")
    
    $first = Get-ADUser -Filter "Enabled -eq 'True' -AND SamAccountName -like 'p*'" -Properties lastLogon,SamAccountName,Department,PasswordLastSet,PasswordNeverExpires,PasswordExpired,EmailAddress | Select-Object Name,SamAccountName,Department,EmailAddress,pass*,@{Name="PasswordAge"; Expression={(Get-Date)-$_.PasswordLastSet}}, @{Name="LastLogon"; Expression={[DateTime]::FromFileTime($_.LastLogon)}} | Sort-Object Name
    $first | Out-GridView -Title "Przygotował Michał Zbyl"
    $TextBox2.AppendText("`r`n")
})
# ping
$Button20.Add_Click({
    $ping = $TextBox3.Text

    if ($ping) {
        foreach ($name in $names){
            if (Test-Connection -ComputerName $name -Count 1 -ErrorAction SilentlyContinue){
                $TextBox2.AppendText("$name,Ping OK`r`n")
            }
            else{
                Write-Host "$name,down"
                $TextBox2.AppendText("$name,Ping notOK`r`n")
            }
        }

        If (Test-Connection $ping -count 1 -quiet) {
                $TextBox2.AppendText("Ping OK`r`n")
        } else {
                $TextBox2.AppendText("Brak łączności`r`n")
        }
    } else {
        $TextBox2.AppendText("Pole puste`r`n")
    }
})
# Ostatni PC
$Button16.Add_Click({
    $TextBox2.AppendText("Czekaj cierpliwie...`r`n")
    $komp = (Get-ADComputer -Filter {Name -Like "K0*"} -Property * | Select-Object -Last 1)
    $kompLast = $komp | Select-Object -ExpandProperty Name
    $TextBox2.AppendText("`r`n")
    $TextBox2.AppendText("Ostatni PC w Domenie to: $kompLast`r`n")
})
# Certyfikat
$Button6.Add_Click({ 
    Clear-Variable $querycheck, $command3, $command, $command2, $wynik, $hasko_Esku, $wynik2, $first, $TextBox3
    $first = $TextBox1.Text
    if ($first) {
        $first = $first.toupper()
        
        ## by SID
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
        $querycheck = "Select * FROM RI_PRACOWNICY where PRAC_USERNAME LIKE '$first'"
        $connection.open()

        $command3=$connection.CreateCommand()
        $command3.CommandText=$querycheck
        $wynik = $command3.ExecuteReader()

        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")

        if ($wynik.HasRows) {
            $query = 'alter user '+$first+' identified by "Szpital.1" account unlock'
            $remad = 'CN='+$first+'_Esk,CN=eskulapConfiguration,DC=SZPITAL,DC=LOCAL'
            $query2 = "update RI_PRACOWNICY SET PRAC_PASS_CHANGE_DATE = (sysdate)-32 WHERE PRAC_USERNAME = '$first'"

            $command=$connection.CreateCommand()
            $command2=$connection.CreateCommand()
            $command.CommandText=$query
            $command2.CommandText=$query2
            $command.ExecuteReader()
            $command2.ExecuteReader()

            Remove-ADObject -Identity $remad -Confirm:$false

            $connection.close()
             
            ## by SID
            $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
            $query = "select *
              from SZ_PRAC_CONFIG
             where PC_ID = ( select max(PC_ID) from SZ_PRAC_CONFIG )"
            
            $connection.open()
            
            $command=$connection.CreateCommand()
            $command.CommandText=$query
            $reader=$command.ExecuteReader()
            
            $table = new-object System.Data.DataTable
            $table.Load($reader)
            
            $connection.close()
            
            $pkiid = $table | Select-Object -ExpandProperty PC_ID
            
            #Write-Output $pkiid
            $pkiidjeden = ($pkiid+1)
            #Write-Output $pkiidjeden
            
            $querycheck = "Select * FROM RI_PRACOWNICY where PRAC_USERNAME LIKE '$first'"
            
            $connection.open()
            
            $command=$connection.CreateCommand()
            $command.CommandText=$querycheck
            $readeruser=$command.ExecuteReader()
            
            $tablepki = new-object System.Data.DataTable
            $tablepki.Load($readeruser)
            
            $connection.close()
            
            $pkiuser = $tablepki | Select-Object -ExpandProperty PRAC_PRACOWNIK_ID
            
            $TextBox2.AppendText("ID użytkownika: $pkiuser`r`n")
            
            $querycheck = "Select * FROM SZ_PRAC_CONFIG where PC_PRAC_PRACOWNIK_ID LIKE '$pkiuser'"
            
            $connection.open()
            
            $command=$connection.CreateCommand()
            $command.CommandText=$querycheck
            $readeruserid=$command.ExecuteReader()
            
            $tablepkiexist = new-object System.Data.DataTable
            $tablepkiexist.Load($readeruserid)
            
            $connection.close()
            
            If ($tablepkiexist.HasRows) {
                $TextBox2.AppendText("Użytkownik posiadał ustawiania logowania kartą w Esku.`r`n")
            } else {
                $queryapkidodaj = "INSERT INTO SZ_PRAC_CONFIG
                (PC_LOW_VALUE, PC_DOMAIN, PC_MEANING, PC_PRAC_PRACOWNIK_ID, PC_INS_USER, PC_JO_JEDNOSTKA_ID, PC_CZY_AKTUALNE, PC_ID)
                VALUES
                ('pki.pass.auto', 'RIPASSWD', 'T', '$pkiuser', 'RI_OWNER', '0', 'T', '$pkiidjeden')"
                $pkiiddwa = ($pkiidjeden+1)

                $connection.open()

                $command=$connection.CreateCommand()
                $command.CommandText=$queryapkidodaj
                $command.ExecuteReader()
        
                $connection.close()

                $queryapkitemp = "INSERT INTO SZ_PRAC_CONFIG
                (PC_LOW_VALUE, PC_DOMAIN, PC_MEANING, PC_PRAC_PRACOWNIK_ID, PC_INS_USER, PC_JO_JEDNOSTKA_ID, PC_CZY_AKTUALNE, PC_ID)
                VALUES
                ('pki.pass.temp', 'RIPASSWD', '0', '$pkiuser', 'RI_OWNER', '0', 'T', '$pkiiddwa')"

                $connection.open()

                $command=$connection.CreateCommand()
                $command.CommandText=$queryapkitemp
                $command.ExecuteReader()
        
                $connection.close()
        
                $TextBox2.AppendText("Dodana możliwość logowania kartą do Eskulapa`r`n")
            }

            $hasko_Esku = "Certyfikat usunięty. Hasło Eskulapa dla $first zostało zmienione na Szpital.1"
        } else {
            $hasko_Esku = "Brak takiego użytkownika: $first"
        } 

        $TextBox2.AppendText("$hasko_Esku")
        $TextBox2.AppendText("`r`n")
    } else {
        $TextBox2.AppendText("Login pusty`r`n")
    }
 })
# Cer-Usuń
$Button13.Add_Click({ 
    Clear-Variable $querycheck, $command3, $command, $command2, $wynik, $hasko_Esku, $wynik2, $first, $TextBox3
    $first = $TextBox1.Text
    if ($first) {
        $first = $first.toupper()
        ## by SID
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
        $querycheck = "Select * FROM RI_PRACOWNICY where PRAC_USERNAME LIKE '$first'"
        $connection.open()

        $command3=$connection.CreateCommand()
        $command3.CommandText=$querycheck
        $wynik = $command3.ExecuteReader()

        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")

        if ($wynik.HasRows) {
            $remad = 'CN='+$first+'_Esk,CN=eskulapConfiguration,DC=SZPITAL,DC=LOCAL'

            Remove-ADObject -Identity $remad -Confirm:$false

            $connection.close()
            ## by SID
            $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
            $query = "select *
              from SZ_PRAC_CONFIG
             where PC_ID = ( select max(PC_ID) from SZ_PRAC_CONFIG )"
            
            $connection.open()
            
            $command=$connection.CreateCommand()
            $command.CommandText=$query
            $reader=$command.ExecuteReader()
            
            $table = new-object System.Data.DataTable
            $table.Load($reader)
            
            $connection.close()
            
            $pkiid = $table | Select-Object -ExpandProperty PC_ID
            
            $querycheck = "Select * FROM RI_PRACOWNICY where PRAC_USERNAME LIKE '$first'"
            
            $connection.open()
            
            $command=$connection.CreateCommand()
            $command.CommandText=$querycheck
            $readeruser=$command.ExecuteReader()
            
            $tablepki = new-object System.Data.DataTable
            $tablepki.Load($readeruser)
            
            $connection.close()
            
            $pkiuser = $tablepki | Select-Object -ExpandProperty PRAC_PRACOWNIK_ID
            
            $TextBox2.AppendText("ID użytkownika: $pkiuser`r`n")
            
            $querycheck = "Select * FROM SZ_PRAC_CONFIG where PC_PRAC_PRACOWNIK_ID LIKE '$pkiuser'"
            
            $connection.open()
            
            $command=$connection.CreateCommand()
            $command.CommandText=$querycheck
            $readeruserid=$command.ExecuteReader()
            
            $tablepkiexist = new-object System.Data.DataTable
            $tablepkiexist.Load($readeruserid)
            
            $connection.close()
            
            If ($tablepkiexist.HasRows) {
                $queryapkidodaj = "DELETE FROM SZ_PRAC_CONFIG WHERE PC_PRAC_PRACOWNIK_ID = '$pkiuser' and PC_LOW_VALUE = 'pki.pass.temp'"

                $connection.open()

                $command=$connection.CreateCommand()
                $command.CommandText=$queryapkidodaj
                $command.ExecuteReader()
        
                $connection.close()

                $queryapkitemp = "DELETE FROM SZ_PRAC_CONFIG WHERE PC_PRAC_PRACOWNIK_ID = '$pkiuser' and PC_LOW_VALUE = 'pki.pass.auto'"

                $connection.open()

                $command=$connection.CreateCommand()
                $command.CommandText=$queryapkitemp
                $command.ExecuteReader()
        
                $connection.close()
            }
            $hasko_Esku = "Usunięta możliwość logowania kartą do Eskulapa"
        } else {
            $hasko_Esku = "Brak takiego użytkownika: $first"
        } 

        $TextBox2.AppendText("$hasko_Esku")
        $TextBox2.AppendText("`r`n")
    } else {
        $TextBox2.AppendText("Login pusty`r`n")
    }
 })
# Konsultacje
$il = 140
$TextBox4.Text = $il
$Button7.Add_Click({ 
    $il = $TextBox4.Text
    $TextBox4.Text = $il
    $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")
    
    ## by SID
    $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
    $query = "select * from (select * FROM OD_KONSULTACJE ORDER BY KON_KONSULTACJA_ID DESC) KON_KONSULTACJA_ID WHERE KON_STATUS NOT LIKE 'B' AND rownum <=$il order by rownum DESC"

    $connection.open()

    $command=$connection.CreateCommand()
    $command.CommandText=$query
    $reader=$command.ExecuteReader()

    $table = new-object System.Data.DataTable
    $table.Load($reader)

    $connection.close()

    $spr = $table | Select-Object KON_KONSULTACJA_ID, KON_DATA, KON_TEKST, KON_TYTUL
    $sprWynik = $spr | Out-GridView -PassThru -Title "Przygotował Michał Zbyl" | select-Object -ExpandProperty 'KON_KONSULTACJA_ID'

    $TextBox2.AppendText("Otworzy się Grid z ostatnimi konsultacjami bez parametru B`r`n")
    $TextBox2.AppendText("OSTATNIE NA DOLE !!! Można sortować`r`n")

    if ($sprWynik) {

        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")
        ## To connect by SID
        $query2 = "update OD_KONSULTACJE set KON_STATUS = 'B', KON_UPD_USER = '', KON_UPD_DATE = '' where KON_KONSULTACJA_ID = $sprWynik"
        $connection.open()

        $command2=$connection.CreateCommand()
        $command2.CommandText=$query2
        $command2.ExecuteReader()

        $connection.close()

        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Konsultacja o ID $sprWynik została odblokowana`r`n")
        $TextBox2.AppendText("`r`n")
    }

 })
# Hasło
$Button1.Add_Click({ 
    Clear-Variable $querycheck, $command3, $command, $command2, $wynik, $hasko_Esku, $wynik2, $first, $TextBox3
    $first = $TextBox1.Text
    if ($first) {
        $first = $first.toupper()
        ## by SID
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
        $querycheck = "Select * FROM RI_PRACOWNICY where PRAC_USERNAME LIKE '$first'"
        $connection.open()

        $command3=$connection.CreateCommand()
        $command3.CommandText=$querycheck
        $wynik = $command3.ExecuteReader()

        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")

        if ($wynik.HasRows) {
            $query = 'alter user '+$first+' identified by "Szpital.1" account unlock'
            $query2 = "update RI_PRACOWNICY SET PRAC_PASS_CHANGE_DATE = (sysdate)-32 WHERE PRAC_USERNAME = '$first'"

            $command=$connection.CreateCommand()
            $command2=$connection.CreateCommand()
            $command.CommandText=$query
            $command2.CommandText=$query2
            $command.ExecuteReader()
            $command2.ExecuteReader()

            $hasko_Esku = "Hasło Eskulapa dla $first zostało zmienione na Szpital.1"
        } else {
            $hasko_Esku = "Brak takiego użytkownika: $first"
        }
        
        $connection.close() 

        $TextBox2.AppendText("$hasko_Esku")
        $TextBox2.AppendText("`r`n")
    } else {
        $TextBox2.AppendText("Login pusty`r`n")
    }
 })
$Button2.Add_Click({ 
    $TextBox2.Text = ''
})
# Status Pacjenta po nazwisku
$Button15.Add_Click({ 
    $TextBox2.AppendText("Piszemy Polskie znaki. Wielkość liter nie ma znaczenia.`r`n")

    $uesku = $TextBox1.Text
    $ueskuf = $uesku.substring(0,1).toupper()+$uesku.substring(1).tolower()

    if ($uesku) {
        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")
        
        ## by SID
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")

        $query = "select p_pacjent_id, p_nazwisko, p_imie, p_nr_pesel, p_status, ps_pobyt_id, po_pobyt_id, ps_data_przyjecia,jo_nazwa, case when p_status = 'A' then 'W domu' when p_status = 'O' then 'Na oddziale' when p_status = 'I' then 'W Izbie Przyjec' when p_status = 'P' then 'W Poradni' when p_status = 'Z' then 'Zgon' else 'Nowy' end from ri_pacjenci, ri_pobyty_w_szpitalu_new, ri_pobyty_na_oddzialach, sz_jednostki_organizacyjne where p_nazwisko = '$ueskuf' and p_pacjent_id = ps_p_pacjent_id and po_ps_pobyt_id = ps_pobyt_id and jo_jednostka_id = po_od_oddzial_id"

        $connection.open()

        $command=$connection.CreateCommand()
        $command.CommandText=$query
        $wynik = $command.ExecuteReader()

        $table = new-object System.Data.DataTable
        $table.Load($wynik)

        $connection.close()

        if ($table) {
            $spr = $table
            $sprll = $spr | Select-Object -Property @{N='ID Pacjenta';E={$_."P_PACJENT_ID"}}, @{N='ID Szpital';E={$_."ps_pobyt_id"}}, @{N='ID Oddział';E={$_."po_pobyt_id"}}, @{N='Data Przyjęcia';E={$_."ps_data_przyjecia"}}, @{N='Imie';E={$_."P_IMIE"}}, @{N='Nazwisko';E={$_."P_NAZWISKO"}}, @{N='Pesel';E={$_."P_NR_PESEL"}}, @{N='Status';E={$_."P_STATUS"}}, @{N='Gdzie';E={$_."CASEWHENP_STATUS='A'THEN'WDOMU'WHENP_STATUS='O'THEN'NAODDZIALE'WHENP_STATUS='I'THEN'WIZBIEPRZYJEC'WHENP_STATUS='P'THEN'WPORADNI'WHENP_STATUS='Z'THEN'ZGON'ELSE'NOWY'END"}}, @{N='Oddział';E={$_."jo_nazwa"}} | Out-GridView -PassThru -Title "Przygotował Michał Zbyl"
            
            $sprel = $sprll
            $TextBox2.AppendText("`r`n")
            $TextBox2.AppendText("ID: $sprel`r`n")
        }
    } else {
        $TextBox2.AppendText("Pole puste`r`n")
    }
})
# MAC/IP
$Button21.Add_Click({ 
    $adreip = $TextBox3.Text
    if ($adreip) {
        $TextBox2.AppendText("To trochę potrwa... Cierpliwości...`r`n")
        if (Get-Module -ListAvailable -Name Posh-SSH) {
            $TextBox2.AppendText("Potrzebne moduły obecne`r`n")
        } else {
            $TextBox2.AppendText("Brak moddułów...Instalacja...`r`n")
            Find-Module Posh-SSH | Install-Module -Confirm:$False -Force
        }
        Import-Module Posh-SSH
        Remove-Item C:\Windows\temp\dhcpd.csv -Force
        Get-SCPFile -RemoteFile "/usr/local/etc/dhcpd.conf" -LocalFile "C:\WINDOWS\TEMP\dhcpd.csv" -ComputerName 192.168.100.2 -Credential (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $l, (ConvertTo-SecureString -String $p -AsPlainText -Force)) -AcceptKey -Force
        
        $line = Get-Content C:\WINDOWS\TEMP\dhcpd.csv | select-string $adreip | Out-GridView -PassThru -Title "Przygotował Michał Zbyl"

        $linef = $line -replace '[:]',''
        $F = $linef -split "(?<=ethernet)\s" | Select-Object -Skip 1 -First 1
        $FF = $F -split "(?<=;)\s" | Select-Object -First 1
        $FFF = $FF -replace '[;]',''

        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("MAC Terminala: $FFF.ini")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Lokalizacja: /var/ftp/wyse/wnos/inc/$FFF.ini")
        $TextBox2.AppendText("`r`n")
        
        Get-SCPFile -RemoteFile "/var/ftp/wyse/wnos/inc/$FFF.ini" -LocalFile "C:\WINDOWS\TEMP\$FFF.ini" -ComputerName 192.168.100.2 -Credential (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $l, (ConvertTo-SecureString -String $p -AsPlainText -Force)) -AcceptKey -Force

        $regex = "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"
        $sel = (Get-Content C:\WINDOWS\TEMP\$FFF.ini | Select-String -Pattern $regex -AllMatches | % { $_.Matches } | % { $_.Value } | 
        ForEach-Object {
            if ($_ -eq "192.168.100.224") {
                Invoke-Item C:\WINDOWS\TEMP\$FFF.ini
                "Drukarka podłączona przez serwer"
                Start-Sleep -s 2
            } else {
                $_
            }
        }) -join "`r`n"

        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Adresy IP Drukarek:`r`n")
        $TextBox2.AppendText("$sel`r`n")

        Remove-SSHSession -Index 0 -Verbose
        Remove-Item C:\Windows\temp\$FFF.ini -Force
        Remove-Item C:\Windows\temp\dhcpd.csv -Force
    } else {
        $TextBox2.AppendText("Brak Szukanego IP/Mac`r`n")
    }
 })
# Email
$Button12.Add_Click({ 
    $ldapu = $TextBox3.Text
    if ($ldapu) {
        $TextBox2.AppendText("Zostanie utworzony email dla użytkownika o podanym p00000`r`n")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("To trochę potrwa... Cierpliwości...`r`n")
        if (Get-Module -ListAvailable -Name Posh-SSH) {
            $TextBox2.AppendText("Potrzebne moduły obecne`r`n")
        } else {
            $TextBox2.AppendText("Brak moddułów...Instalacja...`r`n")
            Find-Module Posh-SSH | Install-Module -Confirm:$False -Force
        }
        Import-Module Posh-SSH
        $session = New-SSHSession -ComputerName 192.168.100.5 -Credential (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $l, (ConvertTo-SecureString -String $p -AsPlainText -Force)) -AcceptKey -Force
        
        $stream = $session.Session.CreateShellStream("PS-SSH", 0, 0, 0, 0, 1000)
        $user = Invoke-SSHCommand $session -Command "whoami"
        $SSHusersName = $user.Output | Out-String
        $SSHusersName = $SSHusersName.Trim()
        $secpas = ConvertTo-SecureString -String $p -AsPlainText -Force
        $results = Invoke-SSHStreamExpectSecureAction -ShellStream $stream -Command "sudo su -" -ExpectString "[sudo] password for $($SSHusersName):" -SecureAction $secpas
        
        $ldapuu = "*"+$ldapu+"*"
        $ldap = (Get-ADUser -Filter {SamAccountName -like $ldapuu} | Select-Object -ExpandProperty DistinguishedName)
        
        $ldapname = (Get-ADUser -Filter {SamAccountName -like $ldapuu} | Select-Object -ExpandProperty Name)
        $ldapnamedot = $ldapname.Replace(" ",".").ToLower()
        
        $Polish = "ą", "ć", "ę", "ł", "ń", "ó", "ś", "ż", "ź"
        $English = "a", "c", "e", "l", "n", "o", "s", "z", "z"
        
        foreach($litera in $ldapnamedot.toCharArray())
        {
            for($i=0; $i -lt 19; $i++)
            {
                if($litera.ToString().Equals($Polish[$i]))
                {
                    $ldapnamedot=$ldapnamedot.Replace($Polish[$i],$English[$i])
                }
            }
        }
        Write-Host = $ldapnamedot
        $ldapnamesplit = $ldapname.Split(" ")
        $ldapnamei = $ldapnamesplit[0]
        $ldapnamen = $ldapnamesplit[1]
        
        $ldapra = "zmprov ra $ldapu@szpital.gorzow.pl $ldapnamedot@szpital.gorzow.pl"
        $ldapmaic = "zmprov ca $ldapnamedot@szpital.gorzow.pl 'P@ssw0rd' cn '$ldapnamei $ldapnamen' sn '$ldapnamen' displayName '$ldapnamei $ldapnamen' givenName '$ldapnamei' description '$ldapu'"
        $ldapma = "zmprov ma $ldapnamedot@szpital.gorzow.pl zimbraAuthLdapExternalDn '$ldap'" 
        $ldapun = "zmprov ma $ldapnamedot@szpital.gorzow.pl zimbraAccountStatus active"
        
        $sReturn = $stream.Read()
        #$stream.WriteLine("whoami")
        Start-Sleep -s 2
        $stream.WriteLine("su zimbra")
        Start-Sleep -s 2
        #$stream.WriteLine("whoami")
        $stream.WriteLine($ldapra)
        $TextBox2.AppendText("$ldapra")
        $TextBox2.AppendText("`r`n")
        Start-Sleep -s 2
        $stream.WriteLine($ldapmaic)
        $TextBox2.AppendText("$ldapmaic")
        $TextBox2.AppendText("`r`n")
        Start-Sleep -s 4
        $stream.WriteLine($ldapma)
        $TextBox2.AppendText("$ldapma")
        $TextBox2.AppendText("`r`n")
        Start-Sleep -s 2
        $stream.WriteLine($ldapun)
        $TextBox2.AppendText("$ldapun")
        $TextBox2.AppendText("`r`n")
        Start-Sleep -s 2
        $sReturn = $stream.Read()
        
        Remove-SSHSession -Index 0 -Verbose
        $TextBox2.AppendText("Gotowe`r`n")
    } else {
        $TextBox2.AppendText("Zostanie utworzony email dla użytkownika o podanym p00000`r`n")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Pole puste`r`n")
    }
 })
# Adres IP
$Button11.Add_Click({ 
    #Start-Process "\\fs01\IT\Raporty\Logon\OstatniRaz.vbs"
    $loginlast = $TextBox3.Text
    if ($loginlast) {
        #$loginlast = "p01161"
        $loginpath = "\\fs01\IT\Raporty\Logon\$loginlast"
        #echo $loginpath

        $loginlastfile = Get-ChildItem $loginpath | Where-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -notLike "RDTS*" } | Sort-Object LastWriteTime | Select-Object -last 1 | Select-Object -ExpandProperty Name
        
        if ($loginlastfile -like 'OstatniRDS.txt') {
            $loginlastfile = 'OstRDSWszystkie.txt'
        }
        
        $loginlastfilefinal = $loginpath+"\" +$loginlastfile
        #echo $loginlastfilefinal

        #echo $lastDataRow

        $lastDataRow = (Get-Content $loginlastfilefinal)[-1]

        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("$loginlast")
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("$lastDataRow")
        $TextBox2.AppendText("`r`n")
        $check_userAD = Get-ADUser -Identity $loginlast | Select-Object -ExpandProperty SurName
        if ($check_userAD) {
            $ListBox.Items.Add("$check_userAD - $loginlast")
        } else {
            $ListBox.Items.Add("$loginlast")
        }
    } else {
        $TextBox2.AppendText("Login Pusty`r`n")
    }
})
$Form.KeyPreview = $True
$Form.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {
	    $Button22.PerformClick()
	}
})
# Połącz
$Button22.Add_Click({ 
    #Start-Process "\\fs01\IT\Raporty\Logon\OstatniRaz.vbs"
    $ippolacz = $TextBox3.Text
    $ipf = $TextBox3.Text

    $loginpath = "\\fs01\IT\Raporty\Logon\$ippolacz"

    $loginlastfile = Get-ChildItem $loginpath | Where-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -notLike "RDTS*" } | Sort-Object LastWriteTime | Select-Object -last 1 | Select-Object -ExpandProperty Name
    
    if ($loginlastfile -like 'OstatniRDS.txt') {
        $loginlastfile = 'OstRDSWszystkie.txt'
    }
    
    $loginlastfilefinal = $loginpath+"\" +$loginlastfile

    $lastDataRow = (Get-Content $loginlastfilefinal)[-1]
    
    $ipl = $lastDataRow.split(' ')[0]
    $check_userAD = Get-ADUser -Identity $ipf | Select-Object -ExpandProperty SurName

    if ($lastDataRow -like '*<========>*') {
        $TextBox3.Text = $ipl
        $ippolacz = $TextBox3.Text
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("$lastDataRow")
        $ListBox.Items.Add("$check_userAD - $ippolacz")
        $TextBox2.AppendText("`r`n")
    } elseif ($lastDataRow -like '*Komputer:*') {
        $iplf = $lastDataRow.split(':')[2]
        $iplff = $iplf.split(' ')[1]
        $TextBox3.Text = $iplff
        $ippolacz = $TextBox3.Text
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("$lastDataRow")
        $ListBox.Items.Add("$check_userAD - $ippolacz")
        $TextBox2.AppendText("`r`n")
    }

    if ($ippolacz) {
        $ippolacz2 = $ippolacz.toupper()
        if ($ippolacz2 -like 'NEG*') {
            $ippolacz = $ippolacz
        } elseif ($ippolacz2 -like 'K*') {
            $ippolacz = $ippolacz
        } elseif ($ippolacz -notlike '192.168*') {
            $ippolacz = '192.168.'+$ippolacz
        }

        Start-Process "C:\Program Files\RealVNC\VNC Viewer\vncviewer.exe" $ippolacz

        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Jeżeli zainstalowany jest C:\Program Files\RealVNC\VNC Viewer`r`n")
        $TextBox2.AppendText("To otworzy się połączenie VNC z $ippolacz")
        $TextBox2.AppendText("`r`n")
        if ($check_userAD) {
            $ListBox.Items.Add("$check_userAD - $ipf")
        } else {
            $ListBox.Items.Add("$ippolacz")
        }
    } else {
        $TextBox2.AppendText("Adres IP pusty`r`n")
    }
})
# Adres IP - Wszystkie
$Button14.Add_Click({
    $loginlast = $TextBox3.Text
    if ($loginlast) {
        $loginpath = "\\fs01\IT\Raporty\Logon\$loginlast"

        Invoke-Item $loginpath

        $TextBox2.AppendText("Explorer otworzony`r`n")
    } else {
        $TextBox2.AppendText("Login Pusty`r`n")
    }
})
# Szukaj loginu po nazwisku - Domena
$Button10.Add_Click({ 
    $TextBox2.AppendText("Piszemy Polskie znaki. Wielkość liter nie ma znaczenia.`r`n")

    $domesku = $TextBox3.Text
    $domeskuf = $domesku.substring(0,1).toupper()+$domesku.substring(1).tolower()

    if ($domeskuf) {
        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")

        $spr = Get-ADUser -Filter "Surname -like '$domeskuf*'" | Select-Object 'Name','SamAccountName'

        if ($spr) {
            $spr3 = $spr | Out-GridView -PassThru -Title "Przygotował Michał Zbyl" | select-Object -ExpandProperty 'SamAccountName'
            $TextBox2.AppendText("`r`n")
            $TextBox2.AppendText("Wynik:`r`n")
            $TextBox2.AppendText("$domesku")
            $TextBox2.AppendText("`r`n")
            $TextBox2.AppendText("$spr3")
            $TextBox2.AppendText("`r`n")
            $TextBox3.Text = $spr3
            $ListBox.Items.Add("$domesku - $spr3")
        } else {
            $TextBox2.AppendText("Brak wyników`r`n")
        }
    } else {
        $TextBox2.AppendText("Pole puste`r`n")
    }
 })
# Szukaj loginu po nazwisku - Esku
$Button9.Add_Click({ 
    $TextBox2.AppendText("Piszemy Polskie znaki. Wielkość liter nie ma znaczenia.`r`n")

    $uesku = $TextBox1.Text
    $ueskuf = $uesku.substring(0,1).toupper()+$uesku.substring(1).tolower()

    if ($uesku) {
        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")
        
        ## by SID
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")

        $query = "Select * FROM RI_PRACOWNICY where PRAC_NAZWISKO LIKE '$ueskuf%'"

        $connection.open()

        $command=$connection.CreateCommand()
        $command.CommandText=$query
        $wynik = $command.ExecuteReader()

        $table = new-object System.Data.DataTable
        $table.Load($wynik)

        $connection.close()

        if ($table) {        
            $spr = $table | Select-Object PRAC_IMIE, PRAC_NAZWISKO, PRAC_USERNAME, PRAC_NR_PESEL, PRAC_PASS_CHANGE_DATE
            $spr3 = $spr | Out-GridView -PassThru -Title "Przygotował Michał Zbyl" | select-Object -ExpandProperty 'PRAC_USERNAME'

            $TextBox2.AppendText("`r`n")
            $TextBox2.AppendText("Wynik:`r`n")
            $TextBox2.AppendText("$uesku")
            $TextBox2.AppendText("`r`n")
            $TextBox2.AppendText("$spr3")
            $TextBox2.AppendText("`r`n")
            $TextBox1.Text = $spr3
        }
    } else {
        $TextBox2.AppendText("Pole puste`r`n")
    }
 })
# Blokada-Wypisy
$Button8.Add_Click({ 
    $TextBox2.AppendText("Lista użytkowników blokujących wypisy.`r`n")
    $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")
     
    ## by SID
    $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
    $locked_object = '$locked_object'
    $session = '$session'
    $query = "select '||a.sid||','||a.serial#||' to_kill, a.username, prac_imie, prac_nazwisko,a.machine, a.lockwait,a.osuser, a.program,b.owner,b.object_name,c.locked_mode from v$locked_object c, all_objects b, v$session a, ri_pracownicy where b.object_id = c.object_id and a.sid = c.session_id and a.username = RI_PRACOWNICY.PRAC_USERNAME and object_name='SZ_WYPISY'"
    
    $connection.open()
    
    $command=$connection.CreateCommand()
    $command.CommandText=$query
    $wynik = $command.ExecuteReader()
    
    $table = new-object System.Data.DataTable
    $table.Load($wynik)
    
    $connection.close()
        
    $spr = $table | Select-Object USERNAME, PRAC_IMIE, PRAC_NAZWISKO, MACHINE, LOCKWAIT, OSUSER, PROGRAM, OWNER
    $spr | Out-GridView -Title "Przygotował Michał Zbyl"
 })
# Odblokuj - Esku
$Button5.Add_Click({ 
    Clear-Variable $querycheck, $command3, $command, $command2, $wynik, $hasko_Esku, $wynik2, $first
    $first = $TextBox1.Text
    if ($first) {
        $first = $first.toupper()
        
        ## by SID
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
        $querycheck = "Select * FROM RI_PRACOWNICY where PRAC_USERNAME LIKE '$first'"
        $connection.open()

        $command3=$connection.CreateCommand()
        $command3.CommandText=$querycheck
        $wynik = $command3.ExecuteReader()

        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")

        if ($wynik.HasRows) {
            $query = 'alter user '+$first+' account unlock'

            $command=$connection.CreateCommand()
            $command.CommandText=$query
            $command.ExecuteReader()

            $hasko_Esku = "$first zostało odblokowane"
        } else {
            $hasko_Esku = "Brak takiego użytkownika: $first"
        }
        
        $connection.close()

        $TextBox2.AppendText("$hasko_Esku")
        $TextBox2.AppendText("`r`n")
    } else {
        $TextBox2.AppendText("Login pusty`r`n")
    }
 })
# Szukaj Nazwiska po p0
$Button23.Add_Click({
    $dom_loginu = $TextBox3.Text
    if ($dom_loginu) {
        $check_userAD = Get-ADUser -Identity $dom_loginu | Select-Object -ExpandProperty Name
        if ($check_userAD) {
            $TextBox2.AppendText("$dom_loginu - $check_userAD`r`n")
        } else {
            $TextBox2.AppendText("Brak takiego loginu $check_userAD")
            $TextBox2.AppendText("`r`n")
        }
    } else {
        $TextBox2.AppendText("Login pusty`r`n")
    }
})
# Odblokuj - Domena
$Button4.Add_Click({ 
    $dom_loginu = $TextBox3.Text
        if ($dom_loginu) {
            $check_userAD = Get-ADUser -Identity $dom_loginu | Select-Object -ExpandProperty Name
            if ($check_userAD) {
                Unlock-ADAccount $dom_loginu
                $TextBox2.AppendText("$dom_loginu ($check_userAD) został odblokowany`r`n")
            } else {
                $TextBox2.AppendText("Brak takiego loginu $dom_loginu")
                $TextBox2.AppendText("`r`n")
            }
    } else {
        $TextBox2.AppendText("Login pusty`r`n")
    }
 })
# Hasło - Domena
$Button3.Add_Click({ 
    $dom_loginf = $TextBox3.Text
    if ($dom_loginf) {
        $check_userAD = Get-ADUser -Identity $dom_loginf | Select-Object -ExpandProperty Name
        $dom_login = $dom_loginf.substring(0,1).tolower()+$dom_loginf.substring(1).tolower()    
        if($dom_login -match "p"){
            $dom_loginl = $dom_login.split('p')[1].split(' ')
            if($dom_loginl -match "^[0-9]*$"){
                if ($dom_login.length -eq 6) {
                    Set-ADAccountPassword $dom_login -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "Szpital.1" -Force -Verbose) -PassThru
                    Unlock-ADAccount $dom_login
                    Set-ADUser -Identity $dom_login -ChangePasswordAtLogon $true
                    $TextBox2.AppendText("Hasło domenowe dla $dom_login ($check_userAD) zostało zmienione na Szpital.1`r`n")
                } else {
                    $TextBox2.AppendText("Musi być 6 znaków 'p00000'`r`n")
                }
            } else {
                $TextBox2.AppendText("Po 'p' muszą być same cyfry`r`n")
            } 
        } else {
            $TextBox2.AppendText("Musi zaczynać sie 'p' (Może być duże)`r`n")
        }
    } else {
        $TextBox2.AppendText("Domenowe hasło bez zmian`r`n")
    }
 })
} Else {
    $TextBox2.AppendText("$user Nie jest IT`r`n")
}

[void]$Form.ShowDialog()