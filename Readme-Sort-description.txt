#######################################
#######################################

#TFGR

#  Alta de usuarios para communities  #####


### PROCESO DESCRIPCION ###
<#
Este Modulo sirve para leer crear usuarios y dar permisos en AD basados en una lista SharePoint donde se recogen los datos.

Se lee la lista
se crea el usuario
Se añade el usuarios a los grupos de AD pertinentes
Se envia usuarios
Se envia contraseña

#>

### ENTRADA DE DATOS ###
<# Los datos los recibiremos de una lista de sharepoint con los siguientes datos. Los datos se chequean en origen


Nombre
Apellidos
Correo Electronico
Telefono Opcional
Empresa del grupo a la que pertenece
Commnunities a las que se le pre-inscribe
País
Estado
                                                 

#>

### REPOS DE git hub
<#

https://github.com/rojorge/AD-Altas-Domain-Module Los ficheros para el controlador de dominio y para el servidor de SharePoint

https://github.com/rojorge/AD-Altas-Domain-Module Modulo de funciones para el controlador de dominio, funciones reutilizables para otros proyectos.


LOs Repos de github se han de copiar en la ruta en C:\Scripts\Communities\

PRINCIPAL
https://github.com/rojorge/AD-Altas-Domain-Module

Contiene scripts a deplegar tanto en el Domain Controller como en el Servidor de SharePoint
Scripts rama AD:
    Process-requests-master: ORCHESTADOR
    Lanzador de process-requestes-master: Es el que ejecuta el la tarea programada y en el se pueden modificar ciertos parametros de forma sencilla, como entorno a ajecutar, pruebas, etc.
Modulo rama AD
    AD-Completa-UserCommunitiesTef: FUnciones propias de AD que son solo validas para el el modulo de altas.

Scripts rama SharePoint:
    SP-leelista-communities: Lee la lista de que relaciona las comunidades con sus valores de grupo de AD, lista de correo asociada, multi, lista de soporte, etc
    SP-leelista-peticiones: Lee la lista de usuarios y devuelve los usuarios aprobados para procesar sus altas.
    Add-membersToCommunities: Lee un archivo xml de objetos tipo SPUSER y añade esos usuarios a una site collection de tipo community, los añade como miembros y actualiza el contador de miembros.
#>
<#
ALTAS AD
Modulo:
    AD-Empresa-Module: FUnciones que se podrá reutilizar para otros proyectos:
        get-LoginFromMail
        get-nameFromEmpresaUSER
        New-SWRandomPassword https://gallery.technet.microsoft.com/scriptcenter/Generate-a-random-and-5c879ed5#content      
        Add-EmpresaUser
        Add-UserToEmpresaGroup
        Send-customEmailFromUser
        add-usersFromGroupToGroup
        
#>

<#
Los ficheros psm1 son modulos de funciones que se añaden al servidor

Los modulos seran desplegados automaticamente en C:\Users\Administrator\Documents\WindowsPowerShell\Modules\"nombre del modulo"
SharePoint:
git@pdihub.hi.inet:ITUSER/AD-Communities.git (SP-leeLista-peticiones, SP-leelista-communities)
git@pdihub.hi.inet:ITUSER/SP-Update-ValueOnList.git
git@pdihub.hi.inet:ITUSER/sp-get-ItemsFromList.git

#>