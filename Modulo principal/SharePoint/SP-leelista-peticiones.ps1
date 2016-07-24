[CmdletBinding(SupportsShouldProcess=$true)]
      param
      (
      
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿Url del site?')]
        $URL,
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿Nombre de la lista?')]
        $nameList,
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿Usuario administrador')]
        $userAdmin

       )

#TFGR

<#
.Synopsis
  Read the sharepoint User list from enrollment and return all the users to add
  
.PREREQUIREMENTS
  C:\Admon\Scripts\CommunitiesBI\sp-get-ItemsFromList\get-ItemsFromList.ps1 must exits
  C:\Scripts\SP13-ALTAS-SharePoint-Functions\SP-Update-ValueOnList\update-ValueOnList.ps1 must exits

.DESCRIPTION
  This script is used to obtain all the user which have been already approved using the generic function get-itemsfromlist.ps1

.EXAMPLE
  c:\Scripts\SP13-ALTAS-SharePoint-Functions\SharePoint\SP-leelista-peticiones.ps1 -URL $web -nameList $nameList -userAdmin $userAdmin} -credential $crd

.OUTPUTS
  [EMPRESA USER]List
.NOTES  
    [EMPRESA USER] is a type of object declared in this script and their attributes have a relation with the internal name of the columns of the user list and are the following:
        Title                                         InternalName                                 
        -----                                         ------------                                 
        Name                                          Title                                        
        Surname                                       apellidos                                    
        Email                                         Email                                        
        Phone Number                                  Telefono                                     
        Company                                       empresa                                      
        Country                                       Pa_x00ed_s                                   
        Communities                                   Community                                    
        Content Type ID                               ContentTypeId                                
        Approver Comments                             _ModerationComments                          
        File Type                                     File_x0020_Type                              
        status                                        estado     


    Written by Jorge Martinez, 
    I take no responsibility for any issues caused by this script.
.FUNCTIONALITY
    Read the sharepoint User list from enrollment and returns all the users approved
    
.REPOSITORY
    https://github.com/rojorge/SP13-AD-Altas
#>

Add-PSSnapin Microsoft.SharePoint.PowerShell -EA SilentlyContinue


####
# Iniciaicialización
$users=@()


### MAIN ###
# Definicion de variables
$web1=$URL
$nameList1=$nameList
$columnTSB="estado"
$valueTSB="Approved (waiting to process)"
$columnToSeekBy2="estado"
$valueTSB2="Approved (waiting to process)"

# Cogemos todos los objetos que cumplen que el estado es "Approved (waiting to process)"

$users=C:Scripts\\sp-get-ItemsFromList\get-ItemsFromList.ps1 -URL $web1 -nameList $nameList1 -columnToSeekby $columnTSB -valueToSeekBy $valueTSB -userAdmin $userAdmin 

# Separamos los usuarios correctos de los que tienen algun problema
$usersToAd=@()
$usersWithErrors=@()

$userWithErrors=@{email="" 
                    error=""
                    }

$emailsUsersWithErrors=@()

ForEach ($user in $users ){
         $usersToAd1=@{
                       Name=$user["Title"]
                       Surname=$user["apellidos"]
                       Email= $user["Email"]
                       PhoneNumber=$user["Telefono"]
                       Company=$user["empresa"]
                       Communities=$user["Community"]
                       Country=$user["Pa_x00ed_s"]
                       #Estado=$user["estado"]
                       Comments=$user["Comentaries"]
                       Login=""
                       Domain=""
                       OU=""
                       MailmanLists=""
                       ADGroupMemberShip=""
                       Password=""
                       ErrorMessage=$user["ErrorMessage"]
                       MacroCommunity=""
                       SupportEmail=""
                       }#userstoad1
                         
        #Actualizamos el estado a "In Process"
        $errorMessage=C:\Scripts\SP13-ALTAS-SharePoint-Functions\SP-Update-ValueOnList\update-ValueOnList.ps1 -URL $URL -nameList $nameList -columnToSeekby "Email" -uniqueValue $user["Email"] -columnToUpdate "estado" -newValue "In Process" -secondaryColumnToSeekBy $columnToSeekBy2 -secondaryUniqueValue $valueTSB2
        
        #si no hay error en la actualizacion añado el usuarios a la variable de usuarios a añadir
        if ([string]::IsNullOrEmpty($errorMessage)) {$usersToAd+=$usersToAd1} 
        else { #si hay error lo añadimos a los usuairos en los que ha habido problema al actualizar el estado.
            if ($emailsUsersWithErrors -notcontains $usersToAd1.email) {$emailsUsersWithErrors+=$usersToAd1.email;$userWithErrors.email=$usersToAd1.email;$userWithErrors.error=$errorMessage;$usersWithErrors+=$userWithErrors}
             }


    }#foreach

#CAmbiamos el estado a los usuairos con error y añaidmos el mensaje de error.
 
$usersWithErrors|% {
            C:\Scripts\SP13-ALTAS-SharePoint-Functions\SP-Update-ValueOnList\update-AllitemsOnList.ps1 -URL $URL -nameList $nameList -columnToSeekby "Email" -uniqueValue $_.email -columnToUpdate "estado" -newValue "error"
            C:\Scripts\SP13-ALTAS-SharePoint-Functions\SP-Update-ValueOnList\update-AllitemsOnList.ps1 -URL $URL -nameList $nameList -columnToSeekby "Email" -uniqueValue $_.email -columnToUpdate "ErrorMessage" -newValue $_.error
            }

return $usersToAd


