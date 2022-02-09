[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$BadAddress = new-object System.Collections.ArrayList
$csv = ''

Function doesUserExist($EmployeeEmail){
    if($EmployeeEmail){
                    $EmployeeEmail = $EmployeeEmail.ToString()
                    $azureaduser = Get-AzureADUser -All $true | Where-Object {$_.Userprincipalname -eq "EmployeeEmail"}
                       #check if something found    
                       if($azureaduser){
                             # Write-Host "User: $UserPrincipalName was found in $displayname AzureAD." -ForegroundColor Green
                             return $true
                             }
                             else{
                             # Write-Host "User $UserPrincipalName was not found in $displayname Azure AD " -ForegroundColor Red
                             $BadAddress.add($EmployeeEmail)
                             return $false
                             }
                        }
}

Function isModuleInstalled($module){
    try{
        Get-InstalledModule -Name $module -ErrorAction Stop
        return $true
    }
    catch [System.Exception]{
          #Write-host "Install Azure AD Module?"
          Install-module AzureAD
          try{
              Get-InstalledModule -Name $module -ErrorAction Stop
              return $true
          }
          catch [System.Exception]{
              return $false
          }

    }
}

Function checkJobTitle($EmployeeEmail){
    $ADJobTitle = (Get-AzureAdUser -ObjectID $EmployeeEmail).jobtitle
    return $ADJobTitle
}

Function checkDepartment($EmployeeEmail){
     $ADDepartment = (Get-AzureAdUser -ObjectID $EmployeeEmail).department
     return $ADDepartment
}

Function checkManager($EmployeeEmail){
    $ADManager = (Get-AzureADUserManager -ObjectID $EmployeeEmail).UserPrincipalName
    return $ADManager
}

Function setJobTitle($JobTitle, $EmployeeEmail){
    Set-AzureADUser -ObjectID $EmployeeEmail -JobTitle $JobTitle
}

Function setDepartment($Department, $EmployeeEmail){
    Set-AzureADUser -ObjectID $EmployeeEmail -Department $Department
}

Function setManager($Manager, $EmployeeEmail){
    Set-AzureADUserManager -ObjectId (Get-AzureADUser -ObjectID $EmployeeEmail).ObjectID  -RefObjectId (Get-AzureADUser -ObjectID $Manager).ObjectID
}

Function ReferenceJobTitle($ADJobTitle, $CSVJobTitle, $EmployeeEmail){
    if ($ADJobTitle -eq [String]::Empty -Or !($ADJobTitle -eq $CSVJobTitle ))
    {
       setJobTitle $CSVJobTitle $EmployeeEmail
    }
}

Function ReferenceDepartment($ADDepartment, $CSVDepartment, $EmployeeEmail){
    if ($ADDepartment -eq [String]::Empty -Or !($ADDepartment -eq $CSVDepartment ))
    {
       setDepartment $CSVDepartment $EmployeeEmail
    }
}

Function ReferenceManager($ADManager, $CSVManager, $EmployeeEmail){
    if ($ADManager -eq [String]::Empty -Or !($ADManager -eq $CSVManager ))
    {
       setManager $CSVManager $EmployeeEmail
    }
}

Function CSVProcess(){
    $csv | ForEach-Object {

    $_
    $EmployeeEmail = $_.'Employee Email'
    $JobTitle = $_.'Job Title'
    $Deparment = $_.Department
    $Manager = $_.'Supervisor Email Address'

        
        if(doesUserExist($EmployeeEmail)){
              $ADJobTitle = checkJobTitle($EmployeeEmail)
              $ADDepartment = checkDepartment($EmployeeEmail)
              $ADManager = checkManager($EmployeeEmail)

              ReferenceJobTitle $ADJobTitle $JobTitle $EmployeeEmail
              ReferenceDepartment $ADDepartment $Department $EmployeeEmail
              ReferenceManager $ADManager $Manager $EmployeeEmail
        }
    }   

}

Function Main(){
    $message=[System.Windows.Forms.MessageBox]::Show("Would you like to process the .csv for updates to Azure AD?","Azure AD CSV Updater",[System.Windows.Forms.MessageBoxButtons]::OKCancel)
     switch ($message){
         "OK" {
             write-host "You pressed OK"
             StartProgram
         }
         "Cancel" {
             write-host "You pressed Cancel"
             # Enter some code
         }
     }
 }


Function getFileName(){
      $File = New-Object System.Windows.Forms.OpenFileDialog
      $File.initialDirectory = [Environment]::GetFolderPath('Desktop')
      $File.filter = "CSV (*.csv)| *.csv"
      $result = $File.ShowDialog()

      if ($result -eq "OK") {
         return  $File.FileName

      }
      return $null
 }

 Function StartProgram(){
     try{
     if(isModuleInstalled('AzureAD')){
         Connect-AzureAD -ErrorAction Stop
         $file = getFileName
         Write-Host "Producing File"
         $file
         if($file)
         {
             Write-Host "In Loop"
             $csv = Import-CSV $file
             CSVProcess
         }
         
         Disconnect-AzureAD
     }
     }
     catch [Microsoft.Open.Azure.AD.CommonLibrary.AadAuthenticationFailedException]{
         Write-Host "Unsuccesful sign in or cancelled sign in, quitting program"
     }
     catch [System.Runtime.InteropServices.COMException]
     {
         Write-Host "Issue with file"
     }

 }
Main
