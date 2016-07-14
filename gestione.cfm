<cfobject component="component" name="comp">

<cfinvoke method="generaExcel" component="#comp#">
    <cfinvokeargument name="stringaInserita" value="#form.txtExcel#">
</cfinvoke>

<cflocation url="index.cfm">