    <body>
        <cfform name="form1" action="gestione.cfm">
            <!--- textarea in cui viene inserito il testo da esportare su excel --->
            <textarea id="txtExcel" name="txtExcel"></textarea>
            <cfinput type="submit" id="btnSubmit" name="btnSubmit" value="Crea file excel"></cfinput>
        </cfform>
    </body>
</html>