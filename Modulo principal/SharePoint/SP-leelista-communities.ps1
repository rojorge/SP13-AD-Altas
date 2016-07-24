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
        HelpMessage='¿Que busco?')]
        $valueToSeekBy,
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='¿Usuario administrador')]
        $userAdmin
       )

#TFGR

<#
.Synopsis
  Read the sharepoint Community-ADGroup-ListaMailman from enrollment and return the values for a specific communty
  
.PREREQUIREMENTS
  C:\Scripts\CommunitiesBI\sp-get-ItemsFromList\get-ItemsFromList.ps1 must exits
.DESCRIPTION
  This script is used to use get-ItemsFromList.ps1 generic script to read the list of Community-ADGroup-ListaMailman

.EXAMPLE
  C:\Admon\Scripts\CommunitiesBI\AD-Communities\SharePoint\SP-leelista-communities.ps1 -URL $URL -nameList $nameList -valueToSeekBy $valueToSeekBy -userAdmin $userAdmin} -credential $Credential

.OUTPUTS
  [SPLIstItem]
.NOTES  
    Written by Jorge Martinez, 
    I take no responsibility for any issues caused by this script.
.FUNCTIONALITY
    Read the sharepoint Community-ADGroup-ListaMailman from enrollment and return the values for a specific communty    
    and returns the looked-for community 
.REPOSITORY
    https://github.com/rojorge/SP13-AD-Altas
#>


Add-PSSnapin Microsoft.SharePoint.PowerShell -EA SilentlyContinue
##




#Invoke-Command -Authentication Credssp -computername $sharepointcomputername -command {param($URL,$nameList,$columnToSeekBy,$valueToSeekBy,$userAdmin) C:\Scripts\SP13-ALTAS-SharePoint-Functions\sp-get-ItemsFromList\get-ItemsFromList.ps1 -URL $URL -nameList $nameList -columnToSeekby $columnToSeekBy -valueToSeekBy $valueToSeekBy -userAdmin $userAdmin} -credential $crd -ArgumentList $URL,$nameList,$columnToSeekBy,$valueToSeekBy,$userAdmin

####
# Iniciaicialización
$objetoCommunities=@()
$columnToSeekBy="Community"


#main


$objetoCommunities=C:\Admon\SP13-ALTAS-SharePoint-Functions\sp-get-ItemsFromList\get-ItemsFromList.ps1 -URL $URL -nameList $nameList -columnToSeekby $columnToSeekBy -valueToSeekBy $valueToSeekBy -userAdmin $userAdmin


#Creamos el objeto comunity con los atributos pertinentes.
$objetoCommunities|%{
         $ObjetoToReturn+=@{
                       Name=$_["Title"]
                       Macro=$_["Macro_x002d_Community"]
                       Mailman=$_["MailMan_x0020_List"]
                       GroupTEF=$_["TEF_x002e_INET_x0020_AdGroup"]
                       GroupMacro=$_["TEF_x002e_INET_x0020_AdGroup_x00"]
                       SupportEmail=$_["e8rc"]
                            } #end objeto


        }#end for each

return $ObjetoToReturn
