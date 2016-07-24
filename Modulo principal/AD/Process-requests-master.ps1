[CmdletBinding(SupportsShouldProcess=$true)]
      param
      (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿Domain, please?')]
        [Boolean]$test,

        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='(PRO,PRE,IC')]
        [String] $entorno
      )

#TFGR

<#
.Synopsis
    This scrip is the orchestator to complete the subscription procees reading from a SHarePoint list
    #ESP Este sera el orquestador del proceso de altas en AD leyendo de una lista sharepoint
.PREREQUIREMENTS
    Fist time it must be run from the console in order to generate the credential files

    WINRM and CREDSSP must be activated in SharePoint Server and Activie Directory server (https://social.technet.microsoft.com/Forums/en-US/1880b61f-851c-4895-acbb-a6cb757b24ee/powershell-remoting-cannot-access-the-local-farm?forum=sharepointadmin)
    A SharePoint list with the proper data must be exist
    Fist time it must be run from the console in order to generate the credential files for a user with rights enough in SharePoint Server


    #ESP Se necesitara activar el winrm y el creedssp tanto en ad como en sharepoint 
    #ESP Debe existir una lista en sharepoint de donde se leeran los datos    
    

.DESCRIPTION
    This scrip is the orchestator to complete the subscription procees reading from a SHarePoint list. This scrip is the onlyone wich will be launched and then it will coordinate all the actions needed.
    THere are the steps:
        Reading the list
        Adding the user to AD
        Adding the user to the reuired AD GROUPS 
        Sendig the user credential
        Sending the password 

    #ESP Este sera el orquestador del proceso de altas en AD leyendo de una lista sharepoint, es decir el script que se lanza y coordina todas las acciones a realizar.
    #ESP Los pasos del proceso son los siguientes:
        #ESP Se lee la lista
        #ESP se crea el usuario
        #ESP Se añade el usuarios a los grupos de AD pertinentes
        #ESP Se envia usuarios
        #ESP Se envia contraseña

.EXAMPLE
    Process-requests-master.ps1 -test $false -entorno "PRO"

.INPUT
   All the data have been checked when were introduced
        Name
        Suraname/s
        Phone Number
        Company
        Country
        Status

    #ESP Los datos los recibiremos de una lista de sharepoint con los siguientes datos. Los datos se chequean en origen
        #ESP Nombre
        #ESP Apellidos
        #ESP Correo Electronico
        #ESP Telefono Opcional
        #ESP Empresa del grupo a la que pertenece
        #ESP Commnunities a las que se le pre-inscribe
        #ESP País
        #ESP Estado
    
.OUTPUTS
    Users should have been processed.
.NOTES
    A type od data has been defined [EMPRESA USER]
    The following list show all the attrinbuttes of this type of dat and the way that they have been feeded:
        Country                        -> Read from SharePoint                                      
        SupportEmail                   get-DetailsFromCommunities
        Name                           -> Read from SharePoint
        ADGroupMemberShip              get-DetailsFromCommunities                                 
        Password                       New-SWRandomPassword
        MailmanLists                   get-DetailsFromCommunities                                               
        Company                        -> Read from SharePoint
        Comments                       -> Read from SharePoint                                              
        MacroCommunity                 get-DetailsFromCommunities                                                
        Domain                         get-domainFromUser
        Macro                          get-DetailsFromCommunities                                                
        ErrorMessage                   -> Read from SharePoint                                              
        Surname                        -> Read from SharePoint                                        
        Estado                         -> Read from SharePoint
        OU                             get-OUFromUser
        PhoneNumber                    -> Read from SharePoint                                              
        Communities                    -> Read from SharePoint
        Email                          -> Read from SharePoint
        Login                          get-LoginFromMail       

    Written by Jorge Martinez, 
    I take no responsibility for any issues caused by this script.
.FUNCTIONALITY
    This scrip is the orchestator to complete the subscription procees reading from a SHarePoint list.
    
.REPOSITORIES
    https://github.com/rojorge/AD-Altas-Domain-Module Files to be installed in SharePoint Server and Active Directory domain controller Server.
    https://github.com/rojorge/AD-Altas-Domain-Module Module to be installed in Active Directory domain controller.
    https://github.com/rojorge/SP13-ALTAS-SharePoint-Functions Generic functions for SharePoint used in this proccess.
    All the repositories must be copied in the following path : c:\scripts\CommunitiesBI\"repo"
#>

  
# Definicion e Inicicialización de las variables dependeindo de los entornos #

switch ($entorno)
    {
"PRE"  {
$web="https://**********"

$nameList="UsersList"
$communitylist="Community-ADGroup-ListaMailman"
$poolADs="ADPoolList"
$sharepointcomputername="**************"
$userAdmin="**********"
$webAdmin="https://************/enrollment/admin/" #/admin
$webCommunities="https://***********/enrollment/admin/"
#$listaSharePointOracle="Communities"
$credFile="crd-sharepoint-prepro.txt"
$pathCredFile="c:\scripts\CommunitiesBI\AD-Communities\"+$credFile
$entorno="PRE"

       }
"PRO" {
$web="https://www.communities.*******.com"

$nameList="UsersList"
$communityList="Community-ADGroup-ListaMailman"
$poolADs="ADPoolList"
$sharepointComputerName="*********"
$userAdmin="*******"
$webAdmin="https://********/enrollment/admin/" #/admin
$webCommunities="https://**********/enrollment/admin/"
#$listaSharePointOracle="Communities"
$credFile="crd-sharepoint-pr.txt"
$pathCredFile="c:\scripts\CommunitiesBI\AD-Communities\"+$credFile
$entorno="PR0"

        }
}


##INicialización resto de las de Variables##

$users=@()
$spuser=@()
[String]$correctUser="YES"
[String]$correctGroup="NO"
$newState="" # se utilizara para saber si el proceso de alta de usuario va bien.
$message=""
$macroCommunity=""


#Actualizamos los ficheros de modulos


cp c:\Scripts\CommunitiesBI\AD-Communities\AD\AD-Completa-UserCommunitiesTef.psm1 -Destination C:\Users\Administrator\Documents\WindowsPowerShell\Modules\AD-Completa-UserCommunitiesTef -force
cp c:\scripts\CommunitiesBI\AD-Empresa-MODULE\AD-Empresa-Module.psm1 -Destination C:\Users\Administrator\Documents\WindowsPowerShell\Modules\AD-Empresa-Module\AD-Empresa-Module.psm1 -force



#se cargban funciones auxiliares #TBD chequear si existe el modulo antes de borrarlo.
if ((get-module "AD-Empresa-Module") -eq $null) {
    Import-Module AD-Empresa-Module
    }
else {
    remove-Module AD-Empresa-Module
    Import-Module AD-Empresa-Module
    }
if ((get-module "AD-Completa-UserCommunitiesTef") -eq $null) {
    Import-Module AD-Completa-UserCommunitiesTef
    }
else {
    remove-Module AD-Completa-UserCommunitiesTef
    Import-Module AD-Completa-UserCommunitiesTef
    }

# Se cargan las credenciales necesarias.

if (-not(Test-Path $pathCredFile)) {Read-Host -AsSecureString | ConvertFrom-SecureString | out-file $pathCredFile}

$pwd = Get-Content $pathCredFile | ConvertTo-SecureString
$crd = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userAdmin,$pwd



### MAIN ###s
#Leemos de la lista sharepoint
#Users almacenara la informacion de los usuarios leida de sharepoint el tipo de objeto es custom y lo definiremos como USUARIO EMPRESA

$users=Invoke-Command -Authentication Credssp -computername $sharepointComputerName -command {param($web,$nameList,$userAdmin) c:\Scripts\CommunitiesBI\AD-Communities\SharePoint\SP-leelista-peticiones.ps1 -URL $web -nameList $nameList -userAdmin $userAdmin} -credential $crd -ArgumentList $webAdmin,$nameList,$userAdmin


write-verbose "lista de sharepoint leida, usuarios a procesar "
write-verbose $users.count 

$usersfull=$users

#completamos los usuarios con el resto de valores de sus atributos *LEER .NOTES

foreach ($spuser in $usersfull) {
    
 #definimos los parametros que necesitamos para hacer el alta
    

    $spuser.domain=get-domainFromUser -email $spuser.email
    $spuser.OU=get-OUFromUser -userDomain $spuser.domain -entorno $entorno -macroCommunity $macroCommunity -verbose
    $Adlistaygrupo=@() #incializamos esta variable que se rellenara con listas de email,los grupos de ad, los nombres de las multi y la cuenta de email de soporte.
    
    if ($spuser.communities -eq $null) {$spuser.communities="ninguna"} # Si no hay community definida, rellenamos el campo con un string que ponga ninguna.
     
# para depuracion con verbose    
    write-verbose "COmmunities"
    write-verbose $spuser.Communities
    write-verbose "URL"
    write-verbose $webCommunities 
    write-verbose "nameList"
    write-verbose $communityList 
    write-verbose "userAdmin"
    write-verbose $userAdmin 
    write-verbose "sharepointcomputername"
    write-verbose $sharepointComputerName
    write-verbose "OU"
    write-verbose $spuser.OU
    write-verbose "credential"
    write-verbose $crd
# FIN de para depuracion con verbose    
    
    #Consultamos lista de correo, grupos de ad, multicommunities e email de soporte y los metemos en la variable $ADlistaygrupo
    $Adlistaygrupo=get-DetailsFromCommunities -communities $spuser.Communities -url $webCommunities -nameList $communityList -userAdmin $userAdmin -sharepointcomputername $sharepointComputerName -Credential $crd
    #Añadimos esos valores al objeto usuario.
    $spuser.MailmanLists=$Adlistaygrupo[0]
    $spuser.ADGroupMemberShip=$Adlistaygrupo[1]
    $spuser.Macro=$Adlistaygrupo[3]
    $spuser.supportEmail=$Adlistaygrupo[2]

    # Añadimos login al objeto usuario
    $spuser.login=get-LoginFromMail -email $spuser.email
    
    # Añadimos login la password al objeto usuario
    $spuser.Password=New-SWRandomPassword -PasswordLength 8 -Count 1

    #capuramos el mensaje de error actual en la varibale $message
    $message=$spuser.errormessage
    $message+=$spuser.login
    
    #actualizamos el mensaje que se añadirá al final a la lista de sharepoint
    $messageFinal=$message
    #actualizamos el campo mensaje de error en la lista de usuarios en sharepoint para este usuario.
    $message=Invoke-Command -Authentication Credssp -computername $sharepointComputerName -command {param($URL,$nameList,$email,$message) c:\Scripts\CommunitiesBI\SP-Update-ValueOnList\update-ValueOnList.ps1 -URL $URL -nameList $nameList -columnToSeekby "Email" -uniqueValue $email -columnToUpdate "ErrorMessage" -newValue $message} -credential $crd -ArgumentList $webAdmin,$nameList,$spuser.email,$message
    $messageFinal+=$message
 
 
 #damos e alta el usuarios
    $correctUser=Add-EmpresaUse -user $spuser
 
    #Actualizamos la variable $messageFinal dependiendo de si el alta ha ido bien o no.
    if ($correctUser -eq $true) { 
        $newState="Done"
        $messageFinal+=" -> Usuario correctamente dado de alta "
        write-verbose "Usuario correctamente dado de alta "
        write-verbose $spuser.login 
        }
    else { 
        write-verbose "alta error "
        write-verbose $spuser.login
        write-verbose $error[0].exception
        $messageFinal+=$error[0].exception 
        $newState="Error"
        

        }#end del error

 #Añadimos el usuario a los grupos de AD pertinentes y actualizamos la variable $mesgeFinal con el valor pertinente dependiendo de si ha ido bien o no.
    if ($correctUser -eq $true){
        $correctGroup=Add-UserToEmpresaGroup -user $spuser
        if ($correctGroup -eq $true) {
         $newState="Done"
         $messageFinal+=" -> Usuario añadido a los grupos de AD "
         write-verbose "Usuario correctamente dado de alta en sus grupos "
         write-verbose $spuser.login
         }
        else { 
         write-verbose "alta grupo error "
         write-verbose $spuser.login
         write-verbose $error[0].exception
         $messageFinal+=" - > alta grupo error "
         $messageFinal+=$error[0].exception
         $newState="Error"
         }#end del error
        }#end del if
    
    #si no se ha defindo que sea un test suponemos que lo es para no enviar correos en las pruebas si se olvida definir le valor $test
    If ($test -eq $null) {$test=$true} 

    #SI ha ido todo bien, enviamos usuario y password.
    if ($newState -eq "Done")
    {
     Send-customEmailFromUser -user $spuser -emailType "Login" -test $Test
     $messageFinal+=" -> Login enviado al usuarios "
     Send-customEmailFromUser -user $spuser -emailType "Password" -test $Test
     $messageFinal+=" -> Password enviada al usuario "
     }#end del if

#actualizamos el estado en la lista SharePoint
    $message=$messageFinal
    
    write-verbose $message
    write-verbose $newState
    

    $lasterror=Invoke-Command -Authentication Credssp -computername $sharepointComputerName -command {param($URL,$nameList,$email,$message) c:\Scripts\CommunitiesBI\SP-Update-ValueOnList\update-ValueOnList.ps1 -URL $URL -nameList $nameList -columnToSeekby "Email" -uniqueValue $email -columnToUpdate "ErrorMessage" -newValue $message -secondaryColumnToSeekBy "estado" -secondaryUniqueValue "In Process"} -credential $crd -ArgumentList $webAdmin,$nameList,$spuser.email,$message
    $lasterror+=Invoke-Command -Authentication Credssp -computername $sharepointComputerName -command {param($URL,$nameList,$email,$newState) c:\Scripts\CommunitiesBI\SP-Update-ValueOnList\update-ValueOnList.ps1 -URL $URL -nameList $nameList -columnToSeekby "Email" -uniqueValue $email -columnToUpdate "estado" -newValue $newState -secondaryColumnToSeekBy "estado" -secondaryUniqueValue "In Process"} -credential $crd -ArgumentList $webAdmin,$nameList,$spuser.email,$newState



   }# % for each

$usersfull #devolvemos todos los usuarios que se han procesado.


