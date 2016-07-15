<!--- oggetto di riferimento al file cfc --->
<cfobject component="component" name="comp">

<!--- controllo sul bottone premuto --->
<cfif StructKeyExists(form, "btnExcel")>
    <!--- definisco la cartella di destinazione del file excel che verrà creato.
    nel caso la funzione GetTempDirectory() non fosse supportata dal browser,
    essa si può sostituire con il path di una cartella a propria scelta --->
    <cfset pathCartellaUtente = "#GetTempDirectory()#">

    <!--- richiamo alla funzione principale per creare il file excel --->
    <cfinvoke method="convertToExcel" component="#comp#">
        <!--- i parametri sono la stringa inserita nella textarea e il path della cartella --->
        <cfinvokeargument name="stringaInserita" value="#form.txtExcel#">
        <cfinvokeargument name="pathCartella" value="#pathCartellaUtente#">
    </cfinvoke>
    
    <!--- reindirizzamento alla pagina di inserimento della stringa --->
    <cflocation url="index.cfm">
</cfif>