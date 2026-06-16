# =====================================================================
#  CREATION AD - AUDIT TOTAL (version anonymisee, tout editable)
#  Onglets calques sur ADUC :
#  General / Compte / Adresse / Organisation / Profil / Membre de /
#  Editeur d'attributs
#
#  Workflow :
#    1) Saisir Prenom / Nom / Login / Modele (DN) / OU cible (DN)
#    2) "Cibler"      -> charge le modele et REMPLIT tous les onglets
#                        + liste les groupes du modele (cochables)
#    3) "Recalculer"  -> recalcule les champs derives (DisplayName, UPN...)
#    4) Editer librement / cocher-decocher / ajouter des groupes
#    5) "Creer"       -> cree l'utilisateur a partir des champs affiches
# =====================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

# =====================================================================
#  CONFIG  ---  CHARGEE DEPUIS UN FICHIER TXT EXTERNE
#  Fichier "Config-AD.txt" attendu dans le meme dossier que le script.
#  Format : Cle = Valeur   ( # = commentaire, lignes vides ignorees )
#  Listes : valeurs separees par des virgules.
#  -> aucune donnee d'entreprise n'est stockee dans ce .ps1
# =====================================================================
$ScriptDir = if ($PSScriptRoot)      { $PSScriptRoot }
             elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
             else                    { (Get-Location).Path }
$ConfigFile = Join-Path $ScriptDir "Config-AD.txt"

# Genere un modele si le fichier n'existe pas, puis arrete
if (-not (Test-Path $ConfigFile)) {
    @"
# =====================================================================
#  Configuration - Creation AD
#  Format : Cle = Valeur   ( # = commentaire )
#  Listes : valeurs separees par des virgules
# =====================================================================
MailDomain             = votre-domaine.tld
LogPath                = C:\Temp\AD_User_Log.csv
DefaultDescription     =
DefaultEmployeeType    =
DefaultExtAttr12Suffix =
DefaultExtAttr13       =
DefaultGroupsToAdd     =
DefaultGroupsToRemove  =
ClearExchangeDefault   = true
"@ | Set-Content -Path $ConfigFile -Encoding ASCII
    [System.Windows.Forms.MessageBox]::Show(
        "Fichier de configuration cree :`n$ConfigFile`n`nRenseigne-le puis relance le script.",
        "Configuration requise") | Out-Null
    return
}

# Lecture cle = valeur (BOM et espaces neutralises)
$cfg = @{}
foreach ($line in (Get-Content -Path $ConfigFile)) {
    $t = $line.Trim().TrimStart([char]0xFEFF)
    if (-not $t -or $t.StartsWith('#')) { continue }
    $i = $t.IndexOf('=')
    if ($i -lt 1) { continue }
    $cfg[$t.Substring(0,$i).Trim()] = $t.Substring($i+1).Trim()
}

$MailDomain             = if ($cfg['MailDomain']) { $cfg['MailDomain'] } else { "votre-domaine.tld" }
$LogPath                = if ($cfg['LogPath'])    { $cfg['LogPath'] }    else { "C:\Temp\AD_User_Log.csv" }
$DefaultDescription     = [string]$cfg['DefaultDescription']
$DefaultEmployeeType    = [string]$cfg['DefaultEmployeeType']
$DefaultExtAttr12Suffix = [string]$cfg['DefaultExtAttr12Suffix']
$DefaultExtAttr13       = [string]$cfg['DefaultExtAttr13']
$DefaultGroupsToAdd     = if ($cfg['DefaultGroupsToAdd']) {
                              $cfg['DefaultGroupsToAdd'] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                          } else { @() }
$DefaultGroupsToRemove  = if ($cfg['DefaultGroupsToRemove']) {
                              $cfg['DefaultGroupsToRemove'] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                          } else { @() }
$ClearExchangeDefault   = ($cfg['ClearExchangeDefault'] -eq 'true')

function New-RandomPassword {
    param ([int]$Length = 12)
    $chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@$%*-_"
    -join (1..$Length | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
}

# =====================================================================
#  FENETRE
# =====================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text          = "Creation AD - Audit total"
$form.Size          = New-Object System.Drawing.Size(620, 720)
$form.StartPosition = "CenterScreen"
$form.Font          = New-Object System.Drawing.Font("Segoe UI", 9)

# ---------- Bandeau du haut : champs pilotes + Cibler ----------
$top = New-Object System.Windows.Forms.GroupBox
$top.Text     = "Identite (pilote les champs calcules)"
$top.Location = New-Object System.Drawing.Point(10, 8)
$top.Size     = New-Object System.Drawing.Size(585, 150)
$form.Controls.Add($top)

function New-TopField {
    param($Label, [int]$Row)
    $l = New-Object System.Windows.Forms.Label
    $l.Text     = $Label
    $l.Location  = New-Object System.Drawing.Point(15, (22 + $Row*28))
    $l.Size      = New-Object System.Drawing.Size(120, 20)
    $top.Controls.Add($l)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location  = New-Object System.Drawing.Point(140, (20 + $Row*28))
    $t.Size      = New-Object System.Drawing.Size(290, 22)
    $top.Controls.Add($t)
    return $t
}
$tbFirst = New-TopField "Prenom"          0
$tbLast  = New-TopField "Nom"             1
$tbLogin = New-TopField "Login"           2
$tbModel = New-TopField "Modele (DN)"     3
$tbOU    = New-TopField "OU cible (DN)"   4

$btnCibler = New-Object System.Windows.Forms.Button
$btnCibler.Text     = "Cibler"
$btnCibler.Location = New-Object System.Drawing.Point(445, 18)
$btnCibler.Size     = New-Object System.Drawing.Size(125, 30)
$top.Controls.Add($btnCibler)

$btnRecalc = New-Object System.Windows.Forms.Button
$btnRecalc.Text     = "Recalculer"
$btnRecalc.Location = New-Object System.Drawing.Point(445, 54)
$btnRecalc.Size     = New-Object System.Drawing.Size(125, 30)
$top.Controls.Add($btnRecalc)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(445, 92)
$lblStatus.Size     = New-Object System.Drawing.Size(125, 50)
$lblStatus.Text     = "Aucun modele charge"
$lblStatus.ForeColor= [System.Drawing.Color]::Gray
$top.Controls.Add($lblStatus)

# =====================================================================
#  ONGLETS
# =====================================================================
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(10, 165)
$tabs.Size     = New-Object System.Drawing.Size(585, 440)
$form.Controls.Add($tabs)

$fields = @{}   # registre cle -> TextBox

function New-Tab {
    param($Title)
    $tp = New-Object System.Windows.Forms.TabPage
    $tp.Text = $Title
    $tabs.TabPages.Add($tp)
    return $tp
}
function Add-Field {
    param($Parent, $Key, $Label, [int]$Row, [bool]$ReadOnly = $false)
    $l = New-Object System.Windows.Forms.Label
    $l.Text     = $Label
    $l.Location  = New-Object System.Drawing.Point(15, (18 + $Row*30))
    $l.Size      = New-Object System.Drawing.Size(170, 20)
    $Parent.Controls.Add($l)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location  = New-Object System.Drawing.Point(190, (15 + $Row*30))
    $t.Size      = New-Object System.Drawing.Size(360, 22)
    $t.ReadOnly  = $ReadOnly
    if ($ReadOnly) { $t.BackColor = [System.Drawing.Color]::FromArgb(240,240,240) }
    $Parent.Controls.Add($t)
    $fields[$Key] = $t
    return $t
}

# ----- Onglet GENERAL -----
$tabGen = New-Tab "General"
Add-Field $tabGen 'GivenName'    "Prenom (GivenName)"   0 | Out-Null
Add-Field $tabGen 'Surname'      "Nom (Surname)"        1 | Out-Null
Add-Field $tabGen 'DisplayName'  "Nom affiche"          2 | Out-Null
Add-Field $tabGen 'Description'  "Description"          3 | Out-Null
Add-Field $tabGen 'Office'       "Bureau (Office)"      4 | Out-Null
Add-Field $tabGen 'EmailAddress' "Email"                5 | Out-Null
Add-Field $tabGen 'HomePage'     "Page web"             6 | Out-Null
$fields['Description'].Text = $DefaultDescription

# ----- Onglet COMPTE -----
$tabAcc = New-Tab "Compte"
Add-Field $tabAcc 'SamAccountName'    "Login (sAMAccountName)" 0 | Out-Null
Add-Field $tabAcc 'UserPrincipalName' "UPN"                    1 | Out-Null
Add-Field $tabAcc 'Password'          "Mot de passe initial"   2 | Out-Null
$fields['Password'].Text = New-RandomPassword

$chkEnabled = New-Object System.Windows.Forms.CheckBox
$chkEnabled.Text     = "Compte active (Enabled)"
$chkEnabled.Checked  = $true
$chkEnabled.Location = New-Object System.Drawing.Point(190, 105)
$chkEnabled.Size     = New-Object System.Drawing.Size(360, 22)
$tabAcc.Controls.Add($chkEnabled)

$chkChangePwd = New-Object System.Windows.Forms.CheckBox
$chkChangePwd.Text     = "Doit changer le mot de passe a l'ouverture"
$chkChangePwd.Checked  = $true
$chkChangePwd.Location = New-Object System.Drawing.Point(190, 132)
$chkChangePwd.Size     = New-Object System.Drawing.Size(360, 22)
$tabAcc.Controls.Add($chkChangePwd)

# ----- Onglet ADRESSE -----
$tabAddr = New-Tab "Adresse"
Add-Field $tabAddr 'StreetAddress' "Rue"            0 | Out-Null
Add-Field $tabAddr 'POBox'         "Boite postale"  1 | Out-Null
Add-Field $tabAddr 'City'          "Ville"          2 | Out-Null
Add-Field $tabAddr 'State'         "Region / Etat"  3 | Out-Null
Add-Field $tabAddr 'PostalCode'    "Code postal"    4 | Out-Null
Add-Field $tabAddr 'Country'       "Pays"           5 | Out-Null

# ----- Onglet ORGANISATION -----
$tabOrg = New-Tab "Organisation"
Add-Field $tabOrg 'Title'      "Fonction (Title)" 0 | Out-Null
Add-Field $tabOrg 'Department' "Service"          1 | Out-Null
Add-Field $tabOrg 'Company'    "Societe"          2 | Out-Null
Add-Field $tabOrg 'Manager'    "Manager (DN)"     3 | Out-Null

# ----- Onglet PROFIL / TELEPHONES -----
$tabProf = New-Tab "Profil / Tel."
Add-Field $tabProf 'HomeDirectory' "Dossier de base"     0 | Out-Null
Add-Field $tabProf 'HomeDrive'     "Lecteur de base"     1 | Out-Null
Add-Field $tabProf 'ProfilePath'   "Chemin du profil"    2 | Out-Null
Add-Field $tabProf 'ScriptPath'    "Script d'ouverture"  3 | Out-Null
Add-Field $tabProf 'OfficePhone'   "Tel. bureau"         4 | Out-Null
Add-Field $tabProf 'MobilePhone'   "Tel. mobile"         5 | Out-Null

# ----- Onglet MEMBRE DE -----
$tabGrp = New-Tab "Membre de"

# --- Section 1 : groupes du modele (a copier) ---
$lblModel = New-Object System.Windows.Forms.Label
$lblModel.Text     = "Groupes du modele (coche = copie sur le nouveau compte) :"
$lblModel.Location = New-Object System.Drawing.Point(12, 8)
$lblModel.Size     = New-Object System.Drawing.Size(555, 18)
$tabGrp.Controls.Add($lblModel)

$clbModel = New-Object System.Windows.Forms.CheckedListBox
$clbModel.Location      = New-Object System.Drawing.Point(12, 28)
$clbModel.Size          = New-Object System.Drawing.Size(553, 140)
$clbModel.CheckOnClick  = $true
$clbModel.DisplayMember = 'Display'
$tabGrp.Controls.Add($clbModel)

$btnAllNone = New-Object System.Windows.Forms.Button
$btnAllNone.Text     = "Tout / Rien"
$btnAllNone.Location = New-Object System.Drawing.Point(12, 172)
$btnAllNone.Size     = New-Object System.Drawing.Size(120, 22)
$tabGrp.Controls.Add($btnAllNone)

# --- Section 2 : groupes a AJOUTER (colonne gauche) ---
$lblAdd = New-Object System.Windows.Forms.Label
$lblAdd.Text      = "Groupes a AJOUTER (coche = ajoute) :"
$lblAdd.Location  = New-Object System.Drawing.Point(12, 200)
$lblAdd.Size      = New-Object System.Drawing.Size(270, 18)
$tabGrp.Controls.Add($lblAdd)

$tbNewGroup = New-Object System.Windows.Forms.TextBox
$tbNewGroup.Location = New-Object System.Drawing.Point(12, 220)
$tbNewGroup.Size     = New-Object System.Drawing.Size(200, 22)
$tabGrp.Controls.Add($tbNewGroup)

$btnAddGroup = New-Object System.Windows.Forms.Button
$btnAddGroup.Text     = "Ajouter"
$btnAddGroup.Location = New-Object System.Drawing.Point(216, 219)
$btnAddGroup.Size     = New-Object System.Drawing.Size(64, 24)
$tabGrp.Controls.Add($btnAddGroup)

$clbAdd = New-Object System.Windows.Forms.CheckedListBox
$clbAdd.Location     = New-Object System.Drawing.Point(12, 248)
$clbAdd.Size         = New-Object System.Drawing.Size(268, 95)
$clbAdd.CheckOnClick = $true
$tabGrp.Controls.Add($clbAdd)
foreach ($g in $DefaultGroupsToAdd) { $clbAdd.Items.Add($g, $true) | Out-Null }

$btnDelGroup = New-Object System.Windows.Forms.Button
$btnDelGroup.Text     = "Retirer la selection"
$btnDelGroup.Location = New-Object System.Drawing.Point(12, 346)
$btnDelGroup.Size     = New-Object System.Drawing.Size(140, 22)
$tabGrp.Controls.Add($btnDelGroup)

# --- Section 3 : groupes a RETIRER (colonne droite) ---
$lblRem = New-Object System.Windows.Forms.Label
$lblRem.Text      = "Groupes a RETIRER (Remove-ADGroupMember) :"
$lblRem.Location  = New-Object System.Drawing.Point(294, 200)
$lblRem.Size      = New-Object System.Drawing.Size(270, 18)
$tabGrp.Controls.Add($lblRem)

$tbRemGroup = New-Object System.Windows.Forms.TextBox
$tbRemGroup.Location = New-Object System.Drawing.Point(294, 220)
$tbRemGroup.Size     = New-Object System.Drawing.Size(200, 22)
$tabGrp.Controls.Add($tbRemGroup)

$btnRemGroup = New-Object System.Windows.Forms.Button
$btnRemGroup.Text     = "Ajouter"
$btnRemGroup.Location = New-Object System.Drawing.Point(498, 219)
$btnRemGroup.Size     = New-Object System.Drawing.Size(64, 24)
$tabGrp.Controls.Add($btnRemGroup)

$clbRem = New-Object System.Windows.Forms.CheckedListBox
$clbRem.Location     = New-Object System.Drawing.Point(294, 248)
$clbRem.Size         = New-Object System.Drawing.Size(268, 95)
$clbRem.CheckOnClick = $true
$tabGrp.Controls.Add($clbRem)
foreach ($g in $DefaultGroupsToRemove) { $clbRem.Items.Add($g, $true) | Out-Null }

$btnDelRem = New-Object System.Windows.Forms.Button
$btnDelRem.Text     = "Retirer la selection"
$btnDelRem.Location = New-Object System.Drawing.Point(294, 346)
$btnDelRem.Size     = New-Object System.Drawing.Size(140, 22)
$tabGrp.Controls.Add($btnDelRem)

# ----- Onglet EDITEUR D'ATTRIBUTS -----
$tabAttr = New-Tab "Editeur d'attributs"
Add-Field $tabAttr 'employeeType'         "employeeType"          0 | Out-Null
Add-Field $tabAttr 'extensionAttribute12' "extensionAttribute12"  1 | Out-Null
Add-Field $tabAttr 'extensionAttribute13' "extensionAttribute13"  2 | Out-Null
Add-Field $tabAttr 'extensionAttribute15' "extensionAttribute15"  3 | Out-Null
$fields['employeeType'].Text         = $DefaultEmployeeType
$fields['extensionAttribute13'].Text = $DefaultExtAttr13

$chkClearExch = New-Object System.Windows.Forms.CheckBox
$chkClearExch.Text     = "Vider les attributs Exchange (legacyExchangeDN, msExch...)"
$chkClearExch.Checked  = $ClearExchangeDefault
$chkClearExch.Location = New-Object System.Drawing.Point(15, 130)
$chkClearExch.Size     = New-Object System.Drawing.Size(540, 40)
$tabAttr.Controls.Add($chkClearExch)

# =====================================================================
#  BAS DE FENETRE : bouton Creer
# =====================================================================
$btnCreer = New-Object System.Windows.Forms.Button
$btnCreer.Text     = "CREER L'UTILISATEUR"
$btnCreer.Location = New-Object System.Drawing.Point(10, 615)
$btnCreer.Size     = New-Object System.Drawing.Size(585, 45)
$btnCreer.BackColor= [System.Drawing.Color]::FromArgb(46,125,50)
$btnCreer.ForeColor= [System.Drawing.Color]::White
$btnCreer.Font     = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnCreer)

# =====================================================================
#  LOGIQUE
# =====================================================================
function Update-Computed {
    $fn = $tbFirst.Text.Trim()
    $ln = $tbLast.Text.Trim()
    $lg = $tbLogin.Text.Trim()
    $ou = $tbOU.Text.Trim()

    $fields['GivenName'].Text         = $fn
    $fields['Surname'].Text           = $ln
    $fields['DisplayName'].Text       = "$ln $fn"
    $fields['SamAccountName'].Text    = $lg
    $fields['UserPrincipalName'].Text = "$lg@$MailDomain"
    $fields['EmailAddress'].Text      = "$fn.$ln@$MailDomain"

    $dn = "CN=$fn $ln,$ou"
    $fields['extensionAttribute12'].Text = "$fn.$ln$DefaultExtAttr12Suffix"
    $fields['extensionAttribute15'].Text = $dn   # provisoire, remplace par le vrai DN a la creation
}

# ---------- CIBLER ----------
$btnCibler.Add_Click({
    if (-not $tbModel.Text.Trim()) {
        [System.Windows.Forms.MessageBox]::Show("Renseigne le DN du modele.") | Out-Null
        return
    }
    try {
        $u = Get-ADUser -Identity $tbModel.Text.Trim() -Properties *
        $Global:SourceUserFull = $u

        # Organisation
        $fields['Title'].Text      = [string]$u.Title
        $fields['Department'].Text = [string]$u.Department
        $fields['Company'].Text    = [string]$u.Company
        $fields['Manager'].Text    = [string]$u.Manager
        # General
        $fields['Office'].Text     = [string]$u.Office
        $fields['HomePage'].Text   = [string]$u.HomePage
        # Adresse
        $fields['StreetAddress'].Text = [string]$u.StreetAddress
        $fields['POBox'].Text         = [string]$u.POBox
        $fields['City'].Text          = [string]$u.City
        $fields['State'].Text         = [string]$u.State
        $fields['PostalCode'].Text    = [string]$u.PostalCode
        $fields['Country'].Text       = [string]$u.Country
        # Profil / Tel
        $fields['HomeDirectory'].Text = [string]$u.HomeDirectory
        $fields['HomeDrive'].Text     = [string]$u.HomeDrive
        $fields['ProfilePath'].Text   = [string]$u.ProfilePath
        $fields['ScriptPath'].Text    = [string]$u.ScriptPath
        $fields['OfficePhone'].Text   = [string]$u.OfficePhone
        $fields['MobilePhone'].Text   = [string]$u.MobilePhone

        # Groupes du modele -> liste cochable (tous coches par defaut)
        $clbModel.Items.Clear()
        foreach ($g in ($u.MemberOf | Sort-Object)) {
            $name = ($g -split ',')[0] -replace '^CN=',''
            $clbModel.Items.Add(
                ([pscustomobject]@{ Display = $name; Value = $g }), $true) | Out-Null
        }

        Update-Computed

        $lblStatus.Text      = "Modele charge ($($clbModel.Items.Count) groupes)"
        $lblStatus.ForeColor = [System.Drawing.Color]::Green
    }
    catch {
        $lblStatus.Text      = "DN invalide"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show("DN invalide : $($_.Exception.Message)") | Out-Null
    }
})

# ---------- RECALCULER ----------
$btnRecalc.Add_Click({ Update-Computed })

# ---------- GESTION DES GROUPES ----------
$btnAllNone.Add_Click({
    $target = -not ($clbModel.CheckedItems.Count -eq $clbModel.Items.Count)
    for ($i=0; $i -lt $clbModel.Items.Count; $i++) { $clbModel.SetItemChecked($i, $target) }
})
$btnAddGroup.Add_Click({
    $g = $tbNewGroup.Text.Trim()
    if ($g) { $clbAdd.Items.Add($g, $true) | Out-Null; $tbNewGroup.Clear() }
})
$tbNewGroup.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter') { $btnAddGroup.PerformClick(); $_.SuppressKeyPress = $true }
})
$btnDelGroup.Add_Click({
    if ($clbAdd.SelectedIndex -ge 0) { $clbAdd.Items.RemoveAt($clbAdd.SelectedIndex) }
})
$btnRemGroup.Add_Click({
    $g = $tbRemGroup.Text.Trim()
    if ($g) { $clbRem.Items.Add($g, $true) | Out-Null; $tbRemGroup.Clear() }
})
$tbRemGroup.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter') { $btnRemGroup.PerformClick(); $_.SuppressKeyPress = $true }
})
$btnDelRem.Add_Click({
    if ($clbRem.SelectedIndex -ge 0) { $clbRem.Items.RemoveAt($clbRem.SelectedIndex) }
})

# ---------- CREER ----------
$btnCreer.Add_Click({
    try {
        if (-not $Global:SourceUserFull) { throw "Cibler un modele d'abord." }

        $fn    = $fields['GivenName'].Text.Trim()
        $ln    = $fields['Surname'].Text.Trim()
        $login = $fields['SamAccountName'].Text.Trim()
        $ou    = $tbOU.Text.Trim()
        if (-not $login) { throw "Login requis." }
        if (-not $ou)    { throw "OU cible requise." }

        $pwdPlain = if ($fields['Password'].Text.Trim()) { $fields['Password'].Text.Trim() } else { New-RandomPassword }
        $pwd      = ConvertTo-SecureString $pwdPlain -AsPlainText -Force

        # --- Parametres obligatoires ---
        $newParams = @{
            Name                  = "$fn $ln"
            GivenName             = $fn
            Surname               = $ln
            DisplayName           = $fields['DisplayName'].Text
            SamAccountName        = $login
            UserPrincipalName     = $fields['UserPrincipalName'].Text
            EmailAddress          = $fields['EmailAddress'].Text
            Path                  = $ou
            AccountPassword       = $pwd
            Enabled               = $chkEnabled.Checked
            ChangePasswordAtLogon = $chkChangePwd.Checked
            Description           = $fields['Description'].Text
        }

        # --- Parametres optionnels (ajoutes seulement si renseignes) ---
        $optMap = @{
            Office='Office'; Company='Company'; Title='Title'; Department='Department'
            StreetAddress='StreetAddress'; POBox='POBox'; City='City'; State='State'
            PostalCode='PostalCode'; HomePage='HomePage'; HomeDirectory='HomeDirectory'
            HomeDrive='HomeDrive'; ProfilePath='ProfilePath'; ScriptPath='ScriptPath'
            OfficePhone='OfficePhone'; MobilePhone='MobilePhone'
        }
        foreach ($p in $optMap.Keys) {
            $v = $fields[$optMap[$p]].Text.Trim()
            if ($v) { $newParams[$p] = $v }
        }

        New-ADUser @newParams

        $UserDN = (Get-ADUser $login).DistinguishedName
        $errors = @()

        # --- Manager ---
        if ($fields['Manager'].Text.Trim()) {
            try { Set-ADUser $login -Manager $fields['Manager'].Text.Trim() }
            catch { $errors += "Manager : $($_.Exception.Message)" }
        }

        # --- Attributs etendus ---
        $replace = @{}
        if ($fields['employeeType'].Text.Trim())         { $replace['employeeType']         = $fields['employeeType'].Text.Trim() }
        if ($fields['extensionAttribute12'].Text.Trim()) { $replace['extensionAttribute12'] = $fields['extensionAttribute12'].Text.Trim() }
        if ($fields['extensionAttribute13'].Text.Trim()) { $replace['extensionAttribute13'] = $fields['extensionAttribute13'].Text.Trim() }
        $ext15 = if ($fields['extensionAttribute15'].Text.Trim()) { $fields['extensionAttribute15'].Text.Trim() } else { $UserDN }
        $replace['extensionAttribute15'] = $ext15
        if ($fields['Country'].Text.Trim()) { $replace['co'] = $fields['Country'].Text.Trim() }
        if ($replace.Count) {
            try { Set-ADUser $login -Replace $replace }
            catch { $errors += "Attributs : $($_.Exception.Message)" }
        }

        # --- Exchange a vider ---
        if ($chkClearExch.Checked) {
            try { Set-ADUser $login -Clear legacyExchangeDN,msExchExtensionCustomAttribute5,msExchHomeServerName }
            catch { $errors += "Exchange clear : $($_.Exception.Message)" }
        }

        # --- Groupes : modele coches + groupes a ajouter coches ---
        $groupsToApply = @()
        foreach ($it in $clbModel.CheckedItems) { $groupsToApply += $it.Value }
        foreach ($it in $clbAdd.CheckedItems)   { $groupsToApply += [string]$it }
        $groupsToApply = $groupsToApply | Where-Object { $_ } | Select-Object -Unique
        foreach ($g in $groupsToApply) {
            try { Add-ADGroupMember -Identity $g -Members $login -ErrorAction Stop }
            catch { $errors += "Groupe '$g' : $($_.Exception.Message)" }
        }

        # --- Groupes a RETIRER (coches) ---
        $groupsToRemove = @()
        foreach ($it in $clbRem.CheckedItems) { $groupsToRemove += [string]$it }
        $groupsToRemove = $groupsToRemove | Where-Object { $_ } | Select-Object -Unique
        foreach ($g in $groupsToRemove) {
            try { Remove-ADGroupMember -Identity $g -Members $login -Confirm:$false -ErrorAction Stop }
            catch { $errors += "Retrait '$g' : $($_.Exception.Message)" }
        }

        # --- Journalisation CSV (preuve de creation) ---
        try {
            $dir = Split-Path $LogPath
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            [pscustomobject]@{
                Date        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                Login       = $login
                DisplayName = $fields['DisplayName'].Text
                UPN         = $fields['UserPrincipalName'].Text
                OU          = $ou
                Modele      = $Global:SourceUserFull.SamAccountName
                Groupes        = ($groupsToApply.Count)
                GroupesRetires = ($groupsToRemove.Count)
                MotDePasse  = $pwdPlain
                Createur    = $env:USERNAME
            } | Export-Csv -Path $LogPath -NoTypeInformation -Append -Encoding UTF8
        } catch { $errors += "Log CSV : $($_.Exception.Message)" }

        if ($errors.Count) {
            [System.Windows.Forms.MessageBox]::Show(
                "Utilisateur cree : $login`r`n`r`nAvertissements :`r`n - " + ($errors -join "`r`n - "),
                "Cree avec avertissements") | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("Utilisateur cree : $login", "Succes") | Out-Null
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur : $($_.Exception.Message)", "Echec") | Out-Null
    }
})

# Recalcul auto quand on tape l'identite
$tbFirst.Add_TextChanged({ Update-Computed })
$tbLast.Add_TextChanged({ Update-Computed })
$tbLogin.Add_TextChanged({ Update-Computed })

[void]$form.ShowDialog()
