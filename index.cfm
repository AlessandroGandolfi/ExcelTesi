    <body>
        <!--- l'attributo enctype="multipart/form-data" permette l'upload dei file --->
        <form name="form1" action="gestione.cfm" method="POST" enctype="multipart/form-data">
            <!--- textarea in cui viene inserito il testo da esportare su excel --->
            <textarea id="txtExcel" name="txtExcel"></textarea>
            <!--- bottone per fare il submit della stringa e avviare la funzione di creazione del file excel --->
            <input type="submit" id="btnExcel" name="btnExcel" value="Crea file excel">
            <!--- input per il selezionamento del file log --->
            <input type="file" name="fileUploadExcel" id="fileUploadExcel" size="30">
        </form>
    </body>
</html>