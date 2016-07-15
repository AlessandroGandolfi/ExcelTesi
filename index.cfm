    <body>
        <cfform name="form1" action="gestione.cfm">
            <!--- textarea in cui viene inserito il testo da esportare su excel --->
            <textarea id="txtExcel" name="txtExcel"></textarea>
            <!--- bottone per fare il submit della stringa e avviare la funzione di creazione del file excel --->
            <cfinput type="submit" id="btnExcel" name="btnExcel" value="Crea file excel"></cfinput>
        </cfform>
    </body>
</html>