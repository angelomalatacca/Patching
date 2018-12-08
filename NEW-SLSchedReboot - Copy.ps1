#REQUIRES -version 2.0

<#
.SYNOPSIS

  NEW-SLSchedReboot - Modulo: SL-Management

.DESCRIPTION

Questo Script schedula il reboot sulle macchine che vengono indicate, prima di schedulare verifica il campo RT nella description in AD.

Se RT = 99:99 NON viene schedulato il reboot e viene RICHIESTO al suo managedby via mail.
Se RT = xx:xx VIENE schedulato il reboot e viene NOTIFICATO al suo managedby via mail.
  
Lo script ha due modalità di esecuzione: Schedulazione e Simulazione. Il default è Simulazione, quindi, se vogliamo schedulare basta specificare il parametro -Schedulazione

Lo script è un wizard con i seguenti step:

1) Viene chiesto di inserire l'utente con cui schedulare il task (Il default è mil\bent).
2) Viene chiesta la password dell'utente.
3) viene chiesto tramite un popup con un calendario la data in cui schedulare il reboot.
4) viene chiesto se vogliamo inviare la mail ai referenti 
5) viene chiesto l'indirizzo mail da cui inviare le notifiche/richieste (Il default è Patching-Sistemi@esselunga.it)
6) Viene chiesto di confermare le informazioni immesse. Y  procede, N  annulla. 

Al termine dello script viene stampato a video un piccolo riepilogo dell'esito ma si consiglia di utilizzare il parametro -OutCsv d:\test.csv per avere l'export dell'esito dettagliato. 

Per sapere come indicare le macchine guardare gli esempi dell'help.

.NOTES

  Author: Angelo Malatacca
  Version: 17.12.05
  Date: 2017-12-05

.EXAMPLE
  NEW-SLSchedReboot -HF_PRODUCTION_A
  
  In questo modo verrà simulata la creazione (e poi cancellato) di un TASK sul gruppo indicato.
  E' possibile anche inviare mail ma NON è consigliabile allarmare i referenti
  
.EXAMPLE
  NEW-SLSchedReboot -HF_PRODUCTION_A -Schedulazione -OutCsv d:\Export.csv
  
  In questo modo verrà schedulato il reboot sul gruppo indicato.
  Tramite il parametro -outcsv verra creato un file CSV con l'esito dettagliato.

.EXAMPLE
  NEW-SLSchedReboot -HF_PRODUCTION_A -Schedulazione -OutCsv d:\Export.csv -SpecialMail
  
  In questo modo verrà schedulato il reboot sul gruppo indicato.
  Tramite il parametro -outcsv verra creato un file CSV con l'esito dettagliato.
  In occasioni particolari, o speciali,  può essere necessario inviare la mail ad una DL impostata nell'extensionAttribute2 in AD.
    
.EXAMPLE
  NEW-SLSchedReboot -HF_PRODUCTION_A -HF_PRODUCTION_B
  
  Lo script può ricevere in input più gruppi.
    
.EXAMPLE
  NEW-SLSchedReboot -ALL
  
  Disabilitato Da codice.
  Progettato per andare su tutto ciò che potrebbe ricevere in input.

 .EXAMPLE
  NEW-SLSchedReboot -HF_PRODUCTION_A  -Only9999  
  
  In questo modo verrà simultata la creazione di un task solo per le macchine del gruppo indicato con RT 99:99. 
  Utile per "sollecitare" i referente ad effettuare un reboot. 
 
.EXAMPLE
  NEW-SLSchedReboot -List d:\temp\ListaServer.txt -schedulazione

  Indicare un file di testo con l'elenco dei server su cui si vuole schedulare un reboot
  
.EXAMPLE
  NEW-SLSchedReboot -ThisGroup "CN=HF-TestReboot,OU=SERVER,DC=MIL,DC=ESSELUNGA,DC=NET"
  
  Oppure
  
  C:\PS>GET-SLRebootTime -ThisGroup "HF-TestReboot"
  
  In questo caso il cmdlet contatterà le macchine presenti nel gruppo indicato. 
#>
[CmdletBinding()]

param(
[Parameter(Mandatory=$false)]
   [Switch] $STAGING1,
 [Parameter(Mandatory=$false)]
   [Switch] $STAGING2,
 [Parameter(Mandatory=$false)]
   [Switch] $STAGING3,
 [Parameter(Mandatory=$false)]
   [Switch] $STAGING4,
 [Parameter(Mandatory=$false)]
   [Switch] $PRODUCTION1,
 [Parameter(Mandatory=$false)]
   [Switch] $PRODUCTION2,
 [Parameter(Mandatory=$false)]
   [Switch] $PRODUCTION3,
 [Parameter(Mandatory=$false)]
   [Switch] $PRODUCTION4,
 [Parameter(Mandatory=$false)]
   [Switch] $TESTReboot,
 [Parameter(Mandatory=$false)]
   [Switch] $Only9999,
 [Parameter(Mandatory=$false)]
   [String] $ThisGroup = $false,
  [Parameter(Mandatory=$false)]
   [String] $OutCSV = $false,
 [Parameter(Mandatory=$false,ValueFromPipeline=$false)]
   [String] $List = $false,
[Parameter(Mandatory=$false)]
   [Switch] $Schedulazione, 
   [Parameter(Mandatory=$false)]
   [Switch] $SpecialMail
)
Begin {
    #Start-Transcript -Path D:\SuperAppoggio\Parente\Patching\Transcript.txt -Force
}
process {

function Main {
        $MachineProcessed = 0 
        $CountErrori = 0
        $CountNonRaggiungibili = 0 
        $CountSchedulati = 0 
        $CountNONSchedulaBILI = 0 
        #Creo un Template per il mio Oggetto
        $objTemplateObject = New-Object psobject
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name ComputerName -Value $null
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name LastReboot -Value $null
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name GapTime -Value $null
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name Description -Value $null
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name Schedulazione -Value $null
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name EsitoSchedulazione -Value $null
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name SistemaOperativo -Value $null
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name ManagedByMail -Value $null
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name MailCC -Value $null    ## Notifica a gruppo DL corrispondende a custom attribute 1 in AD
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name MailCCSpecial -Value $null    ## Notifica a gruppo DL corrispondende a custom attribute 2 in AD da usare in occasioni speciali
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name CN -Value $null
        $objTemplateObject | Add-Member -MemberType NoteProperty -Name Error -Value $null
        $arrDati = @() #Array Di oggetti
        
        # Circolo I membri del gruppo 
        Write-Host "Processing Group: $Choice"
        if ($List -ne $false){  
                $arrCN = $Choice
        }else{
                $arrCN = ([ADSI]"LDAP://$Choice").MEMBER 
        }#End if 
        
        
        
        $arrCN | foreach{
            #Creo Un mio oggetto Temporaneo basandomi sul template 
            $objTemp = $objTemplateObject | Select-Object *
            
            # Estraggo il nome macchina da AD
            $ObjComputer = [ADSI]"LDAP://$_"
            $ComputerName = $ObjComputer.name 
            #se non è specificato only9999 OPPURE se contiene 99 ed è specificato only9999
            $Desc = $ObjComputer.Description
            if (!$Only9999 -OR ($Only9999 -AND ($Desc -like "*99:99*") )){
                Write-Host "Processing Machine: $ComputerName" -BackgroundColor yellow -ForegroundColor black
                $MachineProcessed++
                
                #####    Inizio Recupero Informazioni Generali
                $OraSchedulazione =  out-string -Stream -InputObject $ObjComputer.Description
                $OraSchedulazione = $OraSchedulazione.split(";")
                foreach ($i in 0..($OraSchedulazione.COUNT -1) ) {
                    if($OraSchedulazione[$i].contains("RT=")){
                        $OraSchedulazione = ($OraSchedulazione[$i] -replace "RT=", "")
                        break 
                    }
                }
                $objTemp.schedulazione = $OraSchedulazione
                $objTemp.ComputerName = out-string -Stream -InputObject $ComputerName 
                $Desc =  out-string -Stream -InputObject $ObjComputer.Description
                if ($Desc.contains("D=")){$objTemp.Description = (($Desc.split(";"))[2] -replace "D=", "")}else{$objTemp.Description =   " -- "} 
                $objTemp.SistemaOperativo = out-string -Stream -InputObject $ObjComputer.OperatingSystem
                
                $ManagedBy =  $ObjComputer.ManagedBy
                $objTemp.ManagedByMail =   out-string -Stream -InputObject ([ADSI]"LDAP://$ManagedBy").Mail 
                ########################################################################################
                    # modifica aggiunta mail di gruppo.
                       if ($ObjComputer.extensionAttribute1){$objTemp.MailCC = $ObjComputer.extensionAttribute1 }
                       if ($ObjComputer.extensionAttribute2){$objTemp.MailCCSpecial = $ObjComputer.extensionAttribute2 }
                ########################################################################################
                $objTemp.CN = $_
                
                if (test-connection -Count 1 $ComputerName -Quiet){
                    Write-Host "--->PING OK: $ComputerName"
                    # Estraggo la data dell'ultimo reboot
                   $date = (Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -ErrorVariable err -ErrorAction SilentlyContinue).LastBootUpTime 
                   if ($Err.count -eq 1){
                                $CountErrori++
                                write-host $Err -Foregroundcolor Red
                                Write-Host "Attenzione non si ha i permessi sulla macchina o non è disponibile l'RPC server" -BackgroundColor red -ForegroundColor black 
                                # Scrivo i risultati nell'oggetto temporaneo
                                $objTemp.LastReboot = "NON DEFINIBILE"
                                $objTemp.GapTime =  "NON DEFINIBILE"
                                $objTemp.Error =  out-string -Stream -InputObject $Err
                                #Carico L'oggetto temporaneto all'interno di un singolo Array
                                $arrDati += $objTemp
                        }else{
                                $RebootTime = [System.DateTime]::ParseExact($date.split(".")[0],"yyyyMMddHHmmss",$null) 
                                # Scrivo i risultati nell'oggetto temporaneo
                                $objTemp.LastReboot = $RebootTime
                                    # Calcolo il GAP con oggi per capire da quanto non viene riavviata la macchina
                                    $a = Get-Date $RebootTime #esempio "05/08/2011 09:36:01" 
                                    $c = (Get-Date)  - $a | Select-Object -Property Days, hours, minutes
                                    $k = $c.days.ToString() + " Days " + $c.hours.ToString() + " Hours " + $c.minutes.ToString() + " Minutes" 
                                $objTemp.GapTime =  $k
                                $objTemp.CN = $_
                                $objTemp.Error =  " -- "
                                #Carico L'oggetto temporaneto all'interno di un singolo Array
                                $arrDati += $objTemp
                        } # End if gestione errore
                   #####    Fine Recupero Informazioni Generali     
                   #####    Inizio Schedulazione Reboot
                    
                    
                    #$ComputerName = "w7ictap1v" #Cancellare
                    
              ############################################################################################################ Schedulazione
                    if ($Schedulazione){ # IF Schedulazione / simulazione
                       #################################################################################################################
                       # In questo blocco IF viene SCHEDULATO il reboot
                       if(!$OraSchedulazione.Contains("99:99")){
                           #schtasks /create /tn RestartPatching /tr '"shutdown /r /t 10 /f"' /s $ComputerName /sc once /sd $GiornoSchedulazione /st $OraSchedulazione /ru $UtenteSchedulazione /rp $PasswordUtenteSchedulazione /F | Out-Null
                           
                           Write-Host "Computer: $computername giornoschedulazione: $GiornoSchedulazione O$OraSchedulazione"

                           schtasks /create /tn RestartPatching /tr '"shutdown /r /t 10 /f"' /s $ComputerName /sc once /sd $GiornoSchedulazione /st $OraSchedulazione /ru "SYSTEM" /F | out-null

                            if(!$?){ # Il comando precedente va in errore se non trova il Task 
                                $objTemp.EsitoSchedulazione = "Failed"
                                $CountErrori++
                                Write-Host "ERRORE probabilmente durante la creazione dello scheduled task RestartPatching su $ComputerName"
                                write-host "Comando creazione: $Task"
                                write-host "Comando Verifica: $Task2"
                            }else{
                                $objTemp.EsitoSchedulazione = "OK"
                                $CountSchedulati++
                                Write-Host "--->Creazione Task RestartPatching OK: $ComputerName"
                                SendMail # se non riesce a schedulare non invia la mail. 
                            }
                            # inviare mail NOTIFICA reboot
                        }else{ # Se 99 non schedulo e invio la mail 
                            $objTemp.EsitoSchedulazione = "Non schedulabile"
                            $CountNONSchedulaBILI++
                            # inviare mail RICHIESTA reboot
                            SendMail
                        }
                        #################################################################################################################
                    }else{ # Else Schedulazione / simulazione
                        #################################################################################################################
                        # In questo blocco IF viene effettuata una SIMULAZIONE di schedulazione (creazione e cancellazione del task )
                        if(!$OraSchedulazione.Contains("99:99")){
                            #schtasks /create /tn VerificaSchedulazione /tr  '"TaskdiVerificaSchedulazione"'  /s $ComputerName /sc once /sd $GiornoSchedulazione /st $OraSchedulazione /ru $UtenteSchedulazione /rp $PasswordUtenteSchedulazione /F | out-null
                            schtasks /create /tn VerificaSchedulazione /tr  '"TaskdiVerificaSchedulazione"'  /s $ComputerName /sc once /sd $GiornoSchedulazione /st $OraSchedulazione /ru "SYSTEM" /F | out-null
                            if(!$?){ # Il comando precedente va in errore se non trova il Task 
                                $objTemp.EsitoSchedulazione = "Simulazione - Failed"
                                $CountErrori++
                                Write-Host "ERRORE durante la creazione dello scheduled task VerificaSchedulazione su $ComputerName"
                                write-host "Comando: $Task"
                            }else{
                                $objTemp.EsitoSchedulazione = "Simulazione - OK"
                                $CountSchedulati++
                                Write-Host "--->Creazione Task VerificaSchedulazione OK: $ComputerName"
                                SendMail
                            }
                            schtasks /delete /s $ComputerName /tn VerificaSchedulazione /F | out-null
                            if(!$?){
                                $objTemp.EsitoSchedulazione = "Simulazione - Failed"
                                $CountErrori++
                                WRITE-HOST "ATTENZIONE PROBLEMI DURANTE LA CANCELLAZIONE DEL TASK Di Simulazione"
                            }
                        }else{ #End If !$OraSchedulazione.Contains("99:99")
                            $objTemp.EsitoSchedulazione = "Non schedulabile"
                            $CountNONSchedulaBILI++
                            SendMail
                        }#End if 99
                        #################################################################################################################
                    }# End if Schedulazione / simulazione
                   
                   #####    Fine Schedulazione Reboot
                  }else{ #else test-connection
                        Write-Host "--->PING NO: $ComputerName" -BackgroundColor red -ForegroundColor black
                        # NON PINGA
                        # Scrivo i risultati nell'oggetto temporaneo
                       $objTemp.LastReboot = "NON RAGGIUNGIBILE"
                       $objTemp.GapTime =  "NON RAGGIUNGIBILE"
                       $objTemp.Error =  " -- "
                       $objTemp.EsitoSchedulazione =  " -- "
                       #Carico L'oggetto temporaneto all'interno di un singolo Array
                       $CountNonRaggiungibili++
                       $arrDati += $objTemp
                    }
              }#End if Only 99:99
            
            
            
        }#End Foreach
        #Porto l'oggetto fuori dalla funzione
        
        Write-Host "" 
        Write-Host "Fine Analisi Gruppo, sono state processate $MachineProcessed macchine." -BackgroundColor yellow -ForegroundColor black
        Write-Host "Errori: $CountErrori" 
        Write-Host "Non Raggiungibili: $CountNonRaggiungibili" 
        Write-Host "Schedulati: $CountSchedulati" 
        Write-Host "NON SchedulaBILI (99:99): $CountNONSchedulaBILI"
        Write-Host "====================================================================="
        Write-Host "Press any key to continue ..."
        $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        # Restituisco l'oggetto alla console
        $arrDati 
        return
        
    } #End Function  MAIN

function search-CN{
        #Funzione che cerca le macchine in AD per recuperarne il CN. 
        #Get Domain List
        $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
        $DomainList = @($objForest.Domains | Select-Object Name)
        $Domains = $DomainList | foreach {$_.Name}
        $cnArray = @()

        #Act on each domain
        foreach($Domain in ($Domains))
        {
            Write-Host "Checking $Domain" -fore red
            $ADsPath = [ADSI]"LDAP://$Domain"
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($ADsPath)
            #The filter
            $objSearcher.Filter = "(&(objectCategory=$CosaCercare)(Name=$dacercare))" #
            #$objSearcher.PageSize = 1000
            $objSearcher.SearchScope = "Subtree"
         
            $colResults = $objSearcher.FindAll()
         
            foreach ($objResult in $colResults)
            {
            
                $objArray = $objResult.GetDirectoryEntry()
                $NomeSrv = $objArray.name
                #write-host $objArray.DistinguishedName ";" $objArray.mail ";" $objArray.ProxyAddresses "`r"
                $cnArray += $objArray.DistinguishedName
                
            }
        } 
        $cnArray
        return

}#End function

Function Calendario
{   
     [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
     $WinForm = New-Object Windows.Forms.Form   
     $WinForm.text = "Calendar Control"   
     $WinForm.Size = New-Object Drawing.Size(375,250) 

     $Calendar = New-Object System.Windows.Forms.MonthCalendar 
     $Calendar.MaxSelectionCount = 1     
     $Calendar.SetCalendarDimensions([int]2,[int]1) 
     $WinForm.Controls.Add($Calendar)   
     ############################################################ 
     $btnSched = new-object System.Windows.Forms.Button
     $btnSched.Size = "86,46"
     if ($Schedulazione){$btnSched.Text = "Schedula"}ELSE{$btnSched.Text = "Simula"}
     $btnSched.Location = "1,165"
     $btnSched.add_Click({$WinForm.close()}) 
     $WinForm.Controls.Add($btnSched)
       
     $WinForm.Add_Shown($WinForm.Activate())  
     $WinForm.showdialog() | Out-Null  
     get-date($Calendar.SelectionStart) -uformat "%d/%m/%Y" # Restituisco la data selezionata   get-date -uformat "%d/%m/%Yd"

} #end function Get-DateRange 

Function SendMail{

if ($sendMail -eq 0 ) {
if(!$OraSchedulazione.Contains("99:99")){ # NON Contiene 99:99 quindi NOTIFICO il riavvio al referente

$MessageObject = "NOTIFICA riavvio schedulato del server $ComputerName"
$MessageBody = "E' in corso l'installazione delle patch di sicurezza sui sistemi Windows Server.  
 
Per rendere operative queste correzioni è necessario un restart del server $ComputerName;  sarà effettuato in automatico alle ore $OraSchedulazione di $DataEstesaPerMail.

Hai ricevuto questa mail in quanto referente dei servizi applicativi erogati dal server.

Se così non fosse o per richiedere la modifica dell'orario o del giorno del riavvio contatta il gruppo Sistemisti Open o manda un ticket HDA all'indirizzo sistemi-open@esselunga.it.
Grazie per la collaborazione.

Buona giornata."
            
 }else{ ################################################################################################ Contiene 99:99 quindi chiedo al referente di rebootare 

$MessageObject = "RICHIESTA riavvio server $ComputerName"
$MessageBody = "E' in corso l'installazione delle patch di sicurezza sui sistemi Windows Server.  

Per rendere operative queste correzioni è necessario un restart del server $ComputerName; non abbiamo informazioni sufficienti per effettuare il riavvio in automatico.

Ti chiediamo di provvedere al riavvio del server il giorno $DataEstesaPerMail o subito i giorni successivi.

Hai ricevuto questa mail in quanto referente dei servizi applicativi erogati dal server.
Se così non fosse contatta il gruppo Sistemisti Open o manda un ticket HDA all'indirizzo sistemi-open@esselunga.it.
Grazie per la collaborazione.

Buona giornata."

 } # end if 99:99
    $smtpServer = "smtp.esselunga.net"
    $mailer = New-Object Net.Mail.SMTPclient($smtpServer)
    $emailTo = $objTemp.ManagedByMail
    #$emailTo = "hd@esselunga.it" # Cancellare
    $msg = New-Object Net.Mail.MailMessage($emailFrom,$emailTo,$MessageObject,$MessageBody)
    
    if($objTemp.MailCC){$msg.cc.add($objTemp.MailCC)}   #### NOTIFICA GRUPPO DL

    if($objTemp.MailCCSpecial -and $SpecialMail){
            if ($objTemp.MailCCSpecial -match ';'){
               ($objTemp.MailCCSpecial).split(';') | %{
                    $msg.cc.add($_)
                }
            }else{
                $msg.cc.add($objTemp.MailCCSpecial)
            }
    
    }   #### NOTIFICA GRUPPO DL In occasioni Speciali

    #$msg.Bcc.add("aurelio.parente@esselunga.it")
    $msg.Bcc.add("Patching-Sistemi@esselunga.it")
    $mailer.Send($msg)
}# End if SendMail -eq 0 
}# End function send mail

#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################
######################################################### INIZIO SCRIPT #####################################################################
#############################################################################################################################################
#############################################################################################################################################
#############################################################################################################################################
# Quando si aggiunge un blocco IF per un nuovo Gruppo ricordarsi il parametro ALL (quindi aggiungere il gruppo anche al blocco IF di ALL)     
$risultato = @() # Variabile che conterrà tutto il risultato da esportare in CSV.
$formatoDataUtenteCorrente = (Get-ItemProperty 'HKCU:\Control Panel\International\').sShortDate

## INFORMAZIONI GENERALI
#$UtenteSchedulazione = Read-Host "Inserire l'utente da utilizzare per l'esecuzione del task di reboot (mil\bent)"
if($UtenteSchedulazione -eq ""){$UtenteSchedulazione = "mil\bent"}
#$PasswordUtenteSchedulazione = Read-Host "Inserire la password dell'utente" 
#$GiornoSchedulazione = Calendario
$GiornoSchedulazione = Read-Host "Inserire data schedulazione riavvio $formatoDataUtenteCorrente"
$DataEstesaPerMail = Get-Date $GiornoSchedulazione   -UFormat "%d / %m / %Y
if ($Schedulazione){ $Modalita = "Schedulazione"}else{$Modalita = "Simulazione"} 
# Preparazione prompt YES  NO  invio mail.
write-host "Inviare una mail ai referenti?" -ForegroundColor red -BackgroundColor black
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Invia Mail."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "NON inviare mail."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$sendMail = $host.ui.PromptForChoice($titlePrompt, $message, $options, 1) 
switch ($sendMail)
    {  0{$emailFrom = Read-Host "Mittente Mail (Patching-Sistemi@esselunga.it)"
            if($emailFrom -eq ""){$emailFrom = "Patching-Sistemi@esselunga.it"}
        }# End switch 0 
       1{$emailFrom = " -- "}# End switch 1 
    }
## fINE INFORMAZIONI GENERALI
## RIEPILOGO E ACCETTAZIONE
write-host ""
write-host "Confermi i dati sottostanti?" -ForegroundColor red -BackgroundColor black
$titlePrompt = "Riepilogo:"
$message = "
Giorno Schedulazione: $GiornoSchedulazione
Modalità: $Modalita
User: $UtenteSchedulazione
Password: $PasswordUtenteSchedulazione
Mittente Mail: $emailFrom
Invio Mail: $sendMail

"
 # Preparazione prompt YES  NO  per la conferma dei dati.
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Procedi."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Annulla Tutto."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice($titlePrompt, $message, $options, 1) 
## fINE RIEPILOGO E ACCETTAZIONE
switch ($result)
    {
       0{ # L'utente a cliccato YES
            
            if ($STAGING1){
                            $Choice = "CN=PATCHING-STAGING-1,OU=Patching,OU=SECURITY,DC=MIL,DC=ESSELUNGA,DC=NET"
                            $GruppoXMAil = "PATCHING-STAGING-1"
                            $RisParziale = Main 
                            $risultato += $RisParziale
                         }# END IF PATCHING-STAGING-1
            if ($STAGING2){
                            $Choice = "CN=PATCHING-STAGING-2,OU=Patching,OU=SECURITY,DC=MIL,DC=ESSELUNGA,DC=NET"
                            $GruppoXMAil = "PATCHING-STAGING-2"
                            $RisParziale = Main 
                            $risultato += $RisParziale
                            }# END IF PATCHING-STAGING-2
            if ($STAGING3){
                            $Choice = "CN=PATCHING-STAGING-3,OU=Patching,OU=SECURITY,DC=MIL,DC=ESSELUNGA,DC=NET"
                            $GruppoXMAil = "PATCHING-STAGING-3"
                            $RisParziale = Main 
                            $risultato += $RisParziale
                            }# END IF PATCHING-STAGING-3
            if ($STAGING4){
                            $Choice = "CN=PATCHING-STAGING-4,OU=Patching,OU=SECURITY,DC=MIL,DC=ESSELUNGA,DC=NET"
                            $GruppoXMAil = "PATCHING-STAGING-4"
                            $RisParziale = Main 
                            $risultato += $RisParziale
                            }# END IF PATCHING-STAGING-4
            if ($PRODUCTION1){
                            $Choice = "CN=PATCHING-PROD-1,OU=Patching,OU=SECURITY,DC=MIL,DC=ESSELUNGA,DC=NET"
                            $GruppoXMAil = "PATCHING-PROD-1"
                            $RisParziale = Main 
                            $risultato += $RisParziale
                            }# END IF PATCHING-PROD-1
            if ($PRODUCTION2){
                            $Choice = "CN=PATCHING-PROD-2,OU=Patching,OU=SECURITY,DC=MIL,DC=ESSELUNGA,DC=NET"
                            $GruppoXMAil = "PATCHING-PROD-2"
                            $RisParziale = Main 
                            $risultato += $RisParziale
                            }# END IF PATCHING-PROD-2
            if ($PRODUCTION3){
                            $Choice = "CN=PATCHING-PROD-3,OU=Patching,OU=SECURITY,DC=MIL,DC=ESSELUNGA,DC=NET"
                            $GruppoXMAil = "PATCHING-PROD-3"
                            $RisParziale = Main 
                            $risultato += $RisParziale
                            }# END IF PATCHING-PROD-3
            if ($PRODUCTION4){
                            $Choice = "CN=PATCHING-PROD-4,OU=Patching,OU=SECURITY,DC=MIL,DC=ESSELUNGA,DC=NET"
                            $GruppoXMAil = "PATCHING-PROD-4"
                            $RisParziale = Main 
                            $risultato += $RisParziale
                            }# END IF PATCHING-PROD-4
            if ($TESTReboot){
                            $Choice = "CN=PATCHING-TESTReboot,OU=Patching,OU=SECURITY,DC=MIL,DC=ESSELUNGA,DC=NET"
                            $GruppoXMAil = "PATCHING-TESTReboot"
                            $RisParziale = Main 
                            $risultato += $RisParziale
                            }# END IF PATCHING-PROD-4
            if ($ThisGroup -ne $false){
                            # Se viene specificato il nome del gruppo senza il CN viene cercato il CN su AD con search-CN  
                            if ($thisgroup -like "CN=*"){
                                $Choice = $ThisGroup
                             }else{
                                $CosaCercare = "Group" 
                                $dacercare = $ThisGroup
                                $Choice = search-CN
                             }
                            $GruppoXMAil = $ThisGroup
                            $RisParziale = Main 
                            $risultato += $RisParziale
                            }# END IF $ThisGroup
            if ($List -ne $false){  
                            # Viene ciclato il file | viene Cercato il CN della macchina | viene eseguito MAIN x ogni macchina
                            get-content -Path $List| foreach{
                                        $CosaCercare = "Computer" 
                                        $dacercare = $_
                                        $Choice = search-CN
                                        $GruppoXMAil = "Custom Array"
                                        $RisParziale = Main 
                                        $risultato += $RisParziale
                                }
                            }# End IF List
            if ($OutCSV -ne $false) {$risultato | export-csv $OutCSV -NoTypeInformation} #Esporto In csv
            
            $risultato | Sort-Object -Property LastReboot 
            return
        } # End 0 switch
       1{ # L'utente a cliccato NO
            Write-Host ""
            Write-Host "Hai annullato le seguenti informazioni e terminato lo script." -BackgroundColor black -ForegroundColor red
            Write-Host ""
            Write-Host $message
            
        }# End 1 swintch
     } #End Switch

#Stop-Transcript
}# END PROCESS


