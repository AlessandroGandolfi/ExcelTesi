<!--- oggetto di riferimento al file cfc --->
<cfobject component="component" name="comp">

<!--- controllo sul bottone premuto --->
<cfif StructKeyExists(form, "btnExcel")>
    <!--- definisco la cartella di destinazione del file excel che verrà creato.
    nel caso la funzione GetTempDirectory() non fosse supportata dal browser,
    essa si può sostituire con il path di una cartella a propria scelta --->
    <cfset pathCartellaUtente = "#GetTempDirectory()#">
    <!--- l'upload dei file ha la priorità --->
    <cfif StructKeyExists(form, "fileUploadExcel") and form.fileUploadExcel is not "">
        <cfinvoke method="getFileString" component="#comp#">
            <!--- i parametri sono il file caricato e il path della cartella --->
            <cfinvokeargument name="fullPath" value="#form.fileUploadExcel#">
            <cfinvokeargument name="pathCartella" value="#pathCartellaUtente#">
        </cfinvoke>
        <!--- nel caso non fosse stato eseguito l'upload dei file viene considerato il contenuto della textarea --->
        <cfelseif IsDefined("txtExcel")>
            <!--- richiamo alla funzione principale per creare il file excel --->
            <cfinvoke method="convertToExcel" component="#comp#">
                <!--- i parametri sono la stringa inserita nella textarea e il path della cartella --->
                <cfinvokeargument name="stringaInserita" value="#form.txtExcel#">
                <cfinvokeargument name="pathCartella" value="#pathCartellaUtente#">
            </cfinvoke>
    </cfif>

    <!--- reindirizzamento alla pagina di inserimento della stringa --->
    <cflocation url="index.cfm">
</cfif>