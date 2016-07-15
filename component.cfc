<cfcomponent>
    <!--- funzione per la creazione e download del file excel --->
    <cffunction name="convertToExcel" access="remote">
        <!--- array generato dalla parsificazione della stringa inserita.
        i caratteri chr(13) e chr(10) equivalgono a un tag HTML <br/>, cioè
        un invio a capo --->
        <cfset arrayDatiBase = ListToArray(stringaInserita, chr(13)&chr(10),false,true)>
        
        <!--- matrice che contiene per ogni riga un elemento dell'array precedente,
        le colonne della matrice equivalgono alle colonne del file excel --->
        <cfset matriceDati = ArrayNew(2)>
        <!--- tramite un loop parsifico ogni elemento dell'array e lo salvo in una riga della matrice --->
        <cfloop index="indexLoop" from="1" to="#ArrayLen(arrayDatiBase)#">
            <cfset matriceDati[#indexLoop#] = ListToArray(arrayDatiBase[#indexLoop#], " ", false, false)>
        </cfloop>

        <!--- definizione della query e delle relative colonne del file excel --->
        <cfset queryDati = QueryNew("Date, Time, Date_Time, Version, Server_Startup_Time, Request_ID, Request_Status, CP_Reason, Thread_ID, Client_IP_Address, Request_Method, Request_URL, Execution_Time, Used_Memory, Max_Memory, Used_Memory, Total_Memory, Free_Memory, Query_String, Return_Status_Code, CPU_Time, AMF_Request, JSESSIONID, CFID, CFTOKEN, JDBC_Query_Count, JDBC_Total_Time, JDBC_Total_DB_Time, JDBC_Total_Rows, Bytes_Sent, Time_to_First_Byte, Time_to_Last_Byte, Time_to_Stream_Open, Time_to_Stream_Close, User_Agent_String")>
        
        <!--- array con i nomi delle colonne, queryDati.ColumnList restituisce una lista dei nomi in ordine alfabetico --->
        <cfset arrayNomiColonne = ["Date", "Time", "Date_Time", "Version", "Server_Startup_Time", "Request_ID", "Request_Status", "CP_Reason", "Thread_ID", "Client_IP_Address", "Request_Method", "Request_URL", "Execution_Time", "Used_Memory", "Max_Memory", "Used_Memory", "Total_Memory", "Free_Memory", "Query_String", "Return_Status_Code", "CPU_Time", "AMF_Request", "JSESSIONID", "CFID", "CFTOKEN", "JDBC_Query_Count", "JDBC_Total_Time", "JDBC_Total_DB_Time", "JDBC_Total_Rows", "Bytes_Sent", "Time_to_First_Byte", "Time_to_Last_Byte", "Time_to_Stream_Open", "Time_to_Stream_Close", "User_Agent_String"]>

        <!--- definizione del numero di righe della query, corrispondente al numero di righe della matrice --->
        <cfset righeDati = QueryAddRow(queryDati, ArrayLen(matriceDati))>
        
        <!--- tramite due loop vengono definite le caselle della query, alimentate dai dati della matrice.
        il primo loop fa riferimento al numero di righe della matrice --->
        <cfloop index="indexLoopRiga" from="1" to="#ArrayLen(matriceDati)#">
            <!--- il secondo loop fa riferimento al numero di colonne (caselle) di ogni riga della matrice --->
            <cfloop index="indexLoopColonna" from="1" to="#ArrayLen(matriceDati[indexLoopRiga])#">
                <!--- inserimento dei dati delle caselle della matrice indirizzata dai loop nelle 
                colonne con i nomi presenti nel vettore di nomi delle colonne.
                il primo campo della funzione QuerySetCell() è riferito alla query da alimentare,
                il secondo al nome della colonna, il terzo al dato da inserire nella casella, il quarto al numero della riga --->
                <cfset temp = QuerySetCell(queryDati, arrayNomiColonne[indexLoopColonna], matriceDati[indexLoopRiga][indexLoopColonna], indexLoopRiga)>
            </cfloop>
        </cfloop>

        <!--- stringa contenente il nome del file, formato da "export" seguito da data e ora del momento della creazione del file --->
        <cfset nomeFile = "export_"&DateFormat(Now(), "dd-mm-yyyy")&"_"&TimeFormat(Now(),"HH-nn-ss")&".xls">

        <!--- creazione del file excel riferito alla query e con il nome creati in precedenza --->
        <cfspreadsheet action="write" filename="#pathCartella#\#nomeFile#" query="queryDati" sheetname="#nomeFile#">

        <!--- tag che gestiscono il download del file appena creato --->
        <cfheader name="Content-Disposition" value="attachment; filename=#nomeFile#">  
        <cfcontent type="application/msexcel" file="#pathCartella#\#nomeFile#" deleteFile="yes"> 

    </cffunction>

    <!--- funzione per ricavare la stringa dal file caricato, in seguito richiama la funzione principale --->
    <cffunction name="getFileString" access="remote">
        <!--- salvataggio idel file nella cartella temporanea da cui si può ricavare il full path, 
        per motivi di sicurezza non si può ricavare il path originale del file caricato --->
        <cffile action="upload" filefield="fileUploadExcel" destination="#pathCartella#" nameConflict="overwrite">
        <!--- file caricato e salvataggio del testo all'interno in una stringa --->
        <cfset fileCaricato = "#GetFileFromPath(fullPath)#">
        <cfset stringaLog = "#FileRead(fileCaricato)#">
        <!--- cancello il file da dove ho preso la stringa --->
        <cffile action="delete" file="#fullPath#">
        
        <!--- richiamo della funzione principale, i parametri sono la stringa ricavata dal file e il path della cartella --->
        <cfinvoke method="convertToExcel">
            <cfinvokeargument name="stringaInserita" value="#stringaLog#">
            <cfinvokeargument name="pathCartella" value="#pathCartella#">
        </cfinvoke>
    </cffunction>

</cfcomponent>