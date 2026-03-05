<% @LANGUAGE = "VBSCRIPT" CODEPAGE = "65001" %>
<!--#include virtual="/mrecho/include/top1.asp"-->

<%
db = "omaggio"
tab = "omaggi"
vistoInterno = 0
vistoCopertina = 0
LayoutCopertina = 0
BookTrailer = 0
      Intervista = 0
ComunicatoStampa = 0
ArticoloBlog = 0
PostFacebook = 0
LibroValido = 0

Prevendita = 0

NessunoSconto = 0

eBook = 0
AudioLibro = 0

CorrettoreBozze = 0
AssistenteRedazioneOnline = 0

CampagnaFacebook = 0
LiveStreaming = 0
DigitaleTerrestre = 0
Editing = 0

LavorazioneEsterna = 0

AssunzioneRespon = 0
DocIdentita = 0

Roma2015 = 0
Roma2016 = 0
Roma2017 = 0

Torino2015 = 0
Torino2016 = 0
Torino2017 = 0
Torino2018 = 0
Torino2019 = 0

Milano2017 = 0

Sanremo2013 = 0
Sanremo2014 = 0
Sanremo2015 = 0
Sanremo2016 = 0
Sanremo2019 = 0
Sanremo2020 = 0

redim numeroToRis(5)
numeroToRis(0) = "no"
numeroToRis(1) = "si"
numeroToRis(2) = "si"
numeroToRis(3) = "si"

'percorso = "\\server\stora\bse\autori\"
cn.open connessioneOmaggio()
id = CleanQueryString("id")

aview = request.QueryString("aview")
nascondi_incremento_att = 0

if len(id) > 0 then
    sql = createSelect(tab,top)'"select * from " & tab'
    where = GetSqlValue("AND","id",id,"=")

    query = createSQL(Sql,Where,order,desc)'creazione della query finale con l'orinamento(true = ordinamento decrescente)
    set rs = cn.execute(query)
    id = rs("id")
    nome = rs("nome")
    cognome = rs("cognome")
    indirizzo = rs("indirizzo")
    civico = rs("civico")
    cap = rs("cap")
    citta = rs("citta")
    provincia = rs("provincia")
    tel = rs("tel")
    cell = rs("cell")
    cell2 = rs("cell2")
    cell3 = rs("cell3")
    facebook = rs("facebook")
    email = rs("email")
    if isset(rs("dataIscrizione")) then dataIscrizione = formatDateTime(rs("dataIscrizione"),2)
    if isset(rs("dataInvioCopiaOmaggio")) then dataInvioCopiaOmaggio = formatDateTime(rs("dataInvioCopiaOmaggio"),2)
    pubblicazione = rs("pubblicazione")
    pseudonimo = rs("pseudonimo")
    percorso = rs("percorso")
    vistoInterno = rs("vistoInterno")
    vistoCopertina = rs("vistoCopertina")
    LayoutCopertina = rs("LayoutCopertina")

    InternoPdfVisto = rs("InternoPdfVisto")
    InternoWordRevisione = rs("InternoWordRevisione")
    InternoAvviato = rs("InternoAvviato")

    note = rs("note")
    redattrice = rs ("redattrice")
    grafico = rs("grafico")
    BookTrailer = rs("BookTrailer")
    Intervista = rs("Intervista")
    ComunicatoStampa = rs("ComunicatoStampa")
    ArticoloBlog = rs ("ArticoloBlog")
    PostFacebook = rs ("PostFacebook")
    LibroValido = rs ("LibroValido")

    eBook = rs("eBook")
    AudioLibro = rs("AudioLibro")
    CorrettoreBozze = rs("CorrettoreBozze")
    AssistenteRedazioneOnline = rs("AssistenteRedazioneOnline")

    CampagnaFacebook = rs("CampagnaFacebook")
    LiveStreaming = rs("LiveStreaming")
    DigitaleTerrestre = rs("DigitaleTerrestre")
    Editing = rs("Editing")

    LavorazioneEsterna = rs("LavorazioneEsterna")

    AssunzioneRespon = rs("AssunzioneRespon")
    DocIdentita = rs("DocIdentita")

    Roma2015 = rs("Roma2015")
    Roma2016 = rs("Roma2016")
    Roma2017 = rs("Roma2017")

    Torino2015 = rs("Torino2015")
    Torino2016 = rs("Torino2016")
    Torino2017 = rs("Torino2017")
    Torino2018 = rs("Torino2018")
    Torino2019 = rs("Torino2019")

    Milano2017 = rs("Milano2017")

    Sanremo2013 = rs("Sanremo2013")
    Sanremo2014 = rs("Sanremo2014")
    Sanremo2015 = rs("Sanremo2015")
    Sanremo2016 = rs("Sanremo2016")
    Sanremo2019 = rs("Sanremo2019")
    Sanremo2020 = rs("Sanremo2020")

    Prevendita = rs ("Prevendita")

    NessunoSconto = rs("NessunoSconto")

    noteCommerciale = rs("noteCommerciale")

    predisposizione = rs ("predisposizione")

    provenienza = rs("provenienza")

    if len (provenienza) > 0 then
        provenienza = rs("provenienza")
    else
        provenienza = 0
    end if

    if isset(rs("presso")) then
        presso = rs("presso")
    else
        presso = 0
    end if

    if isset(rs("giorno")) then
        giorno = rs("giorno")
    else
        giorno = 0
    end if

    if isset(rs("mese")) then
        mese = rs("mese")
    else
        mese = 0
    end if

    if isset(rs("anno")) then
        anno = rs("anno")
    else
        anno = 0
    end if

    if isset(rs("sesso")) then
        sesso = rs("sesso")
    else
        sesso = 0
    end if

    if not isset(redattrice) then redattrice = 0

    if not isset(grafico) then grafico = 0

    if isset(rs("provincia")) then
        set cn1 = Server.CreateObject("ADODB.Connection")
        cn1.open connessioneLibrerie()
        set rs = cn1.execute("select * from librerieProvince where id = " & rs("provincia"))
        if not rs.eof then regione = rs("regione")
        cn1.close
    else
        provincia = 0
    end if
else
    sql = createSelect(tab,top)'"select * from " & tab'
end if

cn.close

if not isset(pubblicazione) then pubblicazione = 32
if pubblicazione = 0 then pubblicazione = 32

if not isset(predisposizione) then predisposizione = 1

if not isset(vistoInterno) then vistoInterno = 0
if not isset(vistoCopertina) then vistoCopertina = 0
if not isset(LayoutCopertina) then LayoutCopertina = 0

if not isset(BookTrailer) then BookTrailer = 0
if not isset(Intervista) then Intervista = 0
if not isset(ComunicatoStampa) then ComunicatoStampa = 0
if not isset(ArticoloBlog) then ArticoloBlog = 0
if not isset (PostFacebook) then PostFacebook = 0
if not isset (LibroValido) then LibroValido = 0

if not isset (eBook) then eBook = 0
if not isset (AudioLibro) then AudioLibro = 0
if not isset (CorrettoreBozze) then CorrettoreBozze = 0
if not isset (AssistenteRedazioneOnline) then AssistenteRedazioneOnline = 0

if not isset (CampagnaFacebook) then CampagnaFacebook = 0
if not isset (LiveStreaming) then LiveStreaming = 0
if not isset (DigitaleTerrestre) then DigitaleTerrestre = 0
if not isset (Editing) then Editing = 0

if not isset (LavorazioneEsterna) then LavorazioneEsterna = 0

if not isset (AssunzioneRespon) then AssunzioneRespon = 0
if not isset (DocIdentita) then DocIdentita = 0

if not isset (Roma2015) then Roma2015 = 0
if not isset (Roma2016) then Roma2016 = 0
if not isset (Roma2017) then Roma2017 = 0

if not isset (Torino2015) then Torino2015 = 0
if not isset (Torino2016) then Torino2016 = 0
if not isset (Torino2017) then Torino2017 = 0
if not isset (Torino2018) then Torino2018 = 0
if not isset (Torino2019) then Torino2019 = 0

if not isset (Milano2017) then Milano2017 = 0

if not isset (Sanremo2013) then Sanremo2013 = 0
if not isset (Sanremo2014) then Sanremo2014 = 0
if not isset (Sanremo2015) then Sanremo2015 = 0
if not isset (Sanremo2016) then Sanremo2016 = 0
if not isset (Sanremo2019) then Sanremo2019 = 0
if not isset (Sanremo2020) then Sanremo2020 = 0

if not isset(Prevendita) then Prevendita = 0

if not isset(NessunoSconto) then NessunoSconto = 0

if cInt(NessunoSconto) = 1 then
    nessunoscontoStyle = "background-color:#CC3333; padding-7px; color:#fff; font-weight:bold;"
else
    nessunoscontoStyle = "background-color:#666; color:#fff; font-weight:bold;"
end if

longUrl = CleanQueryString("longUrl")

longUrlSocial = CleanQueryString("longUrlSocial")

'createInputTable(db, sql, nomeCampo, label, valore, richiesto < ! - - - - > )
%>

<div id="titoloArea"><img src="/mrecho/img/frecciaOrizzGialla.gif" width="6" height="11" alt="area di modifica">&nbsp;gestione <%= titolo_pagina %> <a style="float:right; margin-right:104px;" href="./add_tot.asp?id=<%= id %>">Correggi anagrafica</a></div>
<script src="/javas/calendario/js/jscal2.js"></script> 
<script src="/javas/calendario/js/lang/it.js"></script>
<link rel="stylesheet" type="text/css" href="/javas/calendario/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/javas/calendario/css/border-radius.css" />
<script type="text/javascript">
	function visualizza(){
		prendiElementoDaId('add1').style.display = 'block'
		prendiElementoDaId('add2').style.display = 'none'
		prendiElementoDaId('linkx').innerHTML = '<a href=\"javascript:nascondi()\" class=\"linkSelezionato\">[-] Attività svolte</a><a href=\"javascript:visualizza1()\" class=\"linkNormale\">[+] Nuova attività</a>'
		}
	function nascondi(){
		prendiElementoDaId('add1').style.display = 'none'
		prendiElementoDaId('add2').style.display = 'none'
		prendiElementoDaId('linkx').innerHTML = '<a href=\"javascript:visualizza()\" class=\"linkNormale\">[+] Attività svolte</a><a href=\"javascript:visualizza1()\" class=\"linkNormale\">[+] Nuova attività</a>'
		}
	
	function visualizza1(){
		prendiElementoDaId('add2').style.display = 'block'
		prendiElementoDaId('add1').style.display = 'none'
		prendiElementoDaId('linkx').innerHTML = '<a href=\"javascript:visualizza()\" class=\"linkNormale\">[+] Attività svolte</a><a href=\"javascript:nascondi1()\" class=\"linkSelezionato\">[-] Nuova attività</a>'
		}
	function nascondi1(){
		prendiElementoDaId('add1').style.display = 'none'
		prendiElementoDaId('add2').style.display = 'none'
		prendiElementoDaId('linkx').innerHTML = '<a href=\"javascript:visualizza()\" class=\"linkNormale\">[+] Attività svolte</a><a href=\"javascript:visualizza1()\" class=\"linkNormale\">[+] Nuova attività</a>'
		}
	function editAttivita(id){
		window.location = 'add.asp?id=<%= id %>&attivita='+id
	}
	function delAttivita(id){
		window.location = 'eliminaAttivita.asp?sel1='+id
	}
	function setData(svolto){
		today = new Date()
		if(svolto == "1"){
			mese = today.getMonth()+1
			prendiElementoDaId("dataAttivita1").value = today.getDate()+"/"+mese+"/"+today.getFullYear()+" "+today.getHours()+":"+today.getMinutes()
		}else{
			prendiElementoDaId("dataAttivita1").value = ""
		}	
	}
	
	
	
		function setDatafx(giornifx){
		d = new Date()	
	
		date = d.setDate(d.getDate() + giornifx)
		 

/*	date = d.getDate()
      
      
   month = d.getMonth() + 1 // Since getMonth() returns month from 0-11 not 1-12
   year = d.getFullYear()
   hours = d.getHours()
   minutes = d.getMinutes()   
   
   document.getElementById('dataAttivita1').value = date + "/" + month + "/" + year + " " + hours + ":" + minutes
		
*/

document.getElementById('dataAttivita1').value = ("0" + d.getUTCDate()).slice(-2) + "/" +
    ("0" + (d.getUTCMonth()+1)).slice(-2) + "/" +
     d.getUTCFullYear() + " " +
    ("0" + d.getUTCHours()).slice(-2) + ":" +
    ("0" + d.getUTCMinutes()).slice(-2) + ":" +
    ("0" + d.getUTCSeconds()).slice(-2);
     
    	 
	
		}
		
		
function copycopy() {
  const element = document.querySelector('#percorso');
  element.select();
  element.setSelectionRange(0, 99999);
  document.execCommand('copy');
}		
		
		
	
	
	</script>
<%
'CONTROLLO DOPPIONI
set cn3 = Server.CreateObject("ADODB.Connection")
cn3.open connessioneOmaggio()
%>
<form name="add" id="add" action="addConf.asp" method="post" onsubmit="return campiObbligatori(this.name);">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="editTable">
  <tr>
    <td colspan="3" style="background-color:#FFFCE8"><span style="font-weight:bold; color:#000; font-size:13px;"><br/>
      Note sull'autore</span><br/>
      <textarea spellcheck="true" name="facebook"  rows="8" id="facebook" style="width:97%"><%= facebook %></textarea></td>
  </tr>
 
     <%
     if cInt(NessunoSconto) = 1 then
     %>
      <tr><td colspan="2"><span style="display: block;padding:7px;background-color: #CC3333;color:#fff;font-weight: bold;width: 350px;margin: 5px auto;text-align: center;">!!! ATTENZIONE ---> NESSUNO SCONTO !!!</span></td></tr>
      
	  
	  <% end if %>
 
   <%
   if cInt(LavorazioneEsterna) = 1 then
   %>
      <tr><td colspan="2"><span style="display: block;padding:7px;background-color: #e2ff03;color:#585858;font-weight: bold;width: 350px;margin: 5px auto;text-align: center;">!!! ATTENZIONE ---> LAVORAZIONE ESTERNA !!!</span></td></tr>
      
	  
	  <% end if %>
 
 
  <tr>
  
   
      
  
    <%
    if len(id) then

        cn.open connessioneLibri()
        set rsVerificaAutore = cn.execute("select * from autori where idAutoreAttivita=" & id)
        if not rsVerificaAutore.eof then
    %>
    <td colspan="2"><span style="color:#099;font-size:14px;padding:12px;text-align:center;display: block;"><strong>ATTENZIONE:</strong> Ha già pubblicato con noi <br/>
      <%
      set cn8 = Server.CreateObject("ADODB.Connection")
      cn8.open connessioneOmaggio()
      set rs8 = cn8.execute("select * from pubblicazioni order by id asc")
      do while not rs8.eof
          if cInt(pubblicazione) = cInt(rs8("id")) then response.Write(rs8("nome"))
          rs8.movenext
      loop

      cn8.close
      %>
      </span>
     
 
      
      
      </td>
  </tr>
  <tr>
    <td colspan="2"><%
                    i = 0
                    ' set rsLibro = cn.execute("select * from libriQ where autore = " & rsVerificaAutore("id") & " or autore2 = " & rsVerificaAutore("id") & " order by id DESC")

                    set rsLibro = cn.execute("select * from libriQ where idAutoreAttivita = " & id & " or idAutoreAttivita2 = " & id & " order by id DESC")

                    idschedaAutore = rsLibro("autore")

                    do while not rsLibro.eof

                        if cInt(rsLibro("catalogo")) = 1 then
                            valcat = "<span style=""font-size:11px;display: block;margin: 12px 0; color:#6aaf14;"">In Catalogo</span>"
                        else
                            valcat = "<span style=""font-size:11px;display: block;margin: 12px 0; color:#990070;"">Fuori Catalogo</span>"
                        end if
                        if i mod 3 = 0 then response.Write("<div style=""width: 100%; position: relative; float: left; display: inline;"">")
                        set immagineRs = cn.execute("select * from immagini where libri = " & rsLibro("id"))
                        immagine = "nondisp.gif"
                        if not immagineRs.eof then immagine = immagineRs("immagine")
                    %>
      <div style="width: 140px;    padding: 5px; float: left;"><a href="/libro/<%= rsLibro("nome_cat") %>/<%= rsLibro("rewrite") %>" title="<%= rsLibro("titolo") %>" target="_blank"><img src="https://www.booksprint.eu/disk/libri/miniatura/<%= immagine %>" height="195" alt="<%= rsLibro("titolo") %>"></a><br/>
        <% if ( isset(rsLibro("rewrite")) and cInt(rsLibro("catalogo")) = 1 ) then %>
        &nbsp;
        <div id="qrcodeCanvas-<%= rsLibro("rewrite") %>" style="float: left;" class="qrmini"></div>
        
        <!-- Gerena QR CODE --> 
        <script>
		$('#qrcodeCanvas-<%= rsLibro("rewrite") %>').qrcode({
		"text"	:'https://www.booksprintedizioni.it/qr/<%= rsLibro("rewrite") %>',
		"size" 	: 550,
		"render": 'image',
		"quiet" :	1,
		"ecLevel": 'L',
		"background" : '#ffffff',
		"fill" : '#000000'	
		
		
	
	});
</script>
        <% end if %>
        <br/>
        <div style=" font-size: 9px; padding: 0 5px; line-height: 3px; margin-top: -5px; clear: both;"><a style="color:#099;     line-height: 14px; " href="/libro/<%= rsLibro("nome_cat") %>/<%= rsLibro("rewrite") %>" target="_blank">
          <h2><%= rsLibro("titolo") %></h2>
          </a><br />
          <span style="font-size:11px;">Pubblicato: <%= rsLibro("anno") %></span><br/>
          <%= valcat %><br/>
        </div>
        <a style="color: #2C9ADC; font-weight: bold;float:left; text-align:center; display: inline-block;" href="/mrecho/ordini/add.asp?idLibro=<%= rsLibro("id") %>" target="_blank">Avvia Ordine di Stampa</a> <a style="color: #ff8400; font-weight: bold;float:left; text-align:center; display: inline-block;margin-bottom: 20px;" href="/mrecho/libri/add.asp?id=<%= rsLibro("id") %>" target="_blank">Vai a Gestione Libri</a> </div>
      <%
      if i mod 3 = 2 then response.Write("</div>")
      i = i + 1
      rsLibro.movenext
      loop
      if i mod 3 = 2 then response.Write("</div>")
      if i mod 3 = 1 then response.Write("</div>")

      cn.close
      %></td>
    <%
    else
        cn.close
    %>
    <td colspan="2"><span style="color:#099;font-size:18px;padding:12px;text-align:center;display: block;"><strong>Prima Pubblicazione in corso:</strong> <br/>
      <%
      set cn8 = Server.CreateObject("ADODB.Connection")
      cn8.open connessioneOmaggio()
      set rs8 = cn8.execute("select * from pubblicazioni order by id asc")
      do while not rs8.eof
          if cInt(pubblicazione) = cInt(rs8("id")) then response.Write(rs8("nome"))
          rs8.movenext
      loop

      cn8.close
      %>
      </span></td>
    <% end if %>
    <% end if %>
  </tr>
  <tr>
    <td><label for="nome">ID AUTORE Attivit&agrave;: <strong><%= id %></strong></label></td>
    <td><% if len(idschedaAutore) then %>
      <a style="color: #ff8400; font-weight: bold;" href="/mrecho/autori/add.asp?id=<%= idschedaAutore %>" target="_blank">Gestione Pagina Autore Pubblica</a> >>>> ID Autore: <strong><%= idschedaAutore %></strong>
      <%
      set cn8 = Server.CreateObject("ADODB.Connection")
      cn8.open connessioneIscritti()
      set rs8 = cn8.execute("select * from anagrafica where autore=" & idschedaAutore & "")

      if rs8.eof then
      %>
      <br/>
      <span style="display:inline-block; margin:3px; padding:7px; background-color: #CC3333; color:#fff;font-weight: bold;"> <a style="color: #fff; font-weight: bold;" href="/mrecho/iscritti/?q=<%= cognome %>" target="_blank">ACCOUNT ISCRITTO MANCANTE o non configurato</a></span>
      <%
      else

          do while not rs8.eof
      %>
      <br/>
      <a style="color: #00a551; font-weight: bold; display:inline-block; margin:7px 0;" href="/mrecho/iscritti/add.asp?id=<%= rs8("id") %>" target="_blank">Gestione Account Iscritto</a>
      <%
      rs8.movenext
      loop

      end if
      cn8.close
      end if ' chiusura controllo su idschedaAutore
      %></td>
  </tr>
  <tr>
    <td><label for="nome">nome * </label></td>
    <td><input type="text" id="nome" name="nome" value="<%= nome %>" style="width:200px" alt="obbligatorio" maxlength="255">
  </tr>
  <tr>
    <td><label for="cognome">cognome * </label></td>
    <td><input type="text" id="cognome" name="cognome" value="<%= cognome %>" style="width:200px" alt="obbligatorio" maxlength="255">
  </tr>
  <tr>
    <td><label for="indirizzo">indirizzo * </label></td>
    <td><input type="text" id="indirizzo" name="indirizzo" value="<%= indirizzo %>" style="width:200px" alt="obbligatorio" maxlength="255">
  </tr>
  <tr>
    <td><label for="civico">civico</label></td>
    <td><input type="text" id="civico" name="civico" value="<%= civico %>" style="width:200px" maxlength="255">
  </tr>
  <tr>
    <td><label for="cap">cap * </label></td>
    <td><input type="text" id="cap" name="cap" value="<%= cap %>" style="width:200px" alt="obbligatorio" maxlength="255">
  </tr>
  <tr>
    <td><label for="presso">presso </label></td>
    <td><input type="text" id="presso" name="presso" value="<%= presso %>" style="width:200px"  maxlength="255">
  </tr>
  <tr>
    <td><label for="citta">citta * </label></td>
    <td><input type="text" id="citta" name="citta" value="<%= citta %>" style="width:200px" alt="obbligatorio" maxlength="255"></td>
  </tr>
  <tr>
    <td><label for="datadinascita">Data di Nascita </label></td>
    <td><select name="giorno" class="input01 dated" >
        <% if giorno = 0 then %>
        <option value="" disabled selected>giorno</option>
        <% else %>
        <option value="<%= giorno %>" selected="selected"><%= giorno %></option>
        <% end if %>
        <%
        ' CREO LA SELECTBOX PER IL GIORNO
        Dim giorno, d
        giorno = 1
        d = Day(Date())
        For giorno = 1 To 31
        %>
        <option value="<%= giorno %>"><%= giorno %></option>
        <%
        Next
        %>
      </select>
      <%
      ' CREO UN ARRAY PER SPECIFICARE IL GIORNO DELLA SETTIMANA IN FORMATO ALFABETICO
      Dim arr_m(12)
      arr_m(1) = "Gennaio"
      arr_m(2) = "Febbraio"
      arr_m(3) = "Marzo"
      arr_m(4) = "Aprile"
      arr_m(5) = "Maggio"
      arr_m(6) = "Giugno"
      arr_m(7) = "Luglio"
      arr_m(8) = "Agosto"
      arr_m(9) = "Settembre"
      arr_m(10) = "Ottobre"
      arr_m(11) = "Novembre"
      arr_m(12) = "Dicembre"
      %>
      <select name="mese" class="input01 datem" >
        <% if mese = 0 then %>
        <option value="" disabled selected>mese</option>
        <% else %>
        <option value="<%= mese %>" selected="selected"><%= arr_m(mese) %></option>
        <% end if %>
        <%
        ' CREO LA SELECTBOX PER IL MESE
        Dim mese, m
        mese = 1
        m = Month(Date())
        For mese = 1 To UBound(arr_m)
        %>
        <option value="<%= mese %>"><%= arr_m(mese) %></option>
        <%
        Next
        %>
      </select>
      <select name="anno" class="input01 datey" >
        <% if anno = 0 then %>
        <option value="" disabled selected>anno</option>
        <% else %>
        <option value="<%= anno %>" selected="selected"><%= anno %></option>
        <% end if %>
        <%
        Dim anno
        anno = 1
        For anno = 1900 To Year(Date())
        %>
        <option value="<%= anno %>"><%= anno %></option>
        <%
        Next
        %>
      </select></td>
  </tr>
  <tr>
  <td>Sesso </td>
  <td>
  <select name="sesso">
        <option value="0" <% if cInt(sesso) = 0 then response.Write("selected") %>>NON SELEZIONATO</option>
        <option value="1" <% if cInt(sesso) = 1 then response.Write("selected") %>>DONNA</option>
        <option value="2" <% if cInt(sesso) = 2 then response.Write("selected") %>>UOMO</option>
        <option value="3" <% if cInt(sesso) = 3 then response.Write("selected") %>>ALTRO</option>
      
               
      </select></td>
  </tr>
  
  
  
  <tr>
    <td>Nazione </td>
    <td><%
        if isset(provincia) then
            cn.open connessioneLibrerie()
            set rs = cn.execute("select * from librerieProvince where id = " & provincia)
            if not rs.eof then
                set rsNazione = cn.execute("select * from nazioni where id = " & rs("stato"))
                response.Write(rsNazione("nome"))
            end if
            cn.close
        else
            response.Write("Non selezionato")
        end if
        %></td>
  </tr>
  <tr>
    <td><label for="regione">Regione * </label></td>
    <td><select name="regione" onChange="makeRequest('province.asp?regione='+this[this.selectedIndex].value,'provinceCont');prendiElementoDaId('provControllo').value = ''">
        <option value="">Seleziona la regione...</option>
        <%
        cn.open connessioneLibrerie()
        set rs = cn.execute("select distinct regione from librerieProvince order by regione asc")
        do while not rs.eof
            response.Write("<option value=""" & rs(0) & """ ")
            if regione = rs(0) then response.Write(" selected")
            response.Write(">" & rs(0) & "</option>")
            rs.movenext
        loop
        cn.close
        %>
      </select></td>
  </tr>
  <tr>
    <td><label for="prov">Provincia * </label></td>
    <td><span id="provinceCont">
      <select name="prov" onChange="prendiElementoDaId('provincia').value = this[this.selectedIndex].value">
        <%
        cn.open connessioneLibrerie()
        set rs = cn.execute("select * from librerieProvince where regione = '" & replace(regione,"'","''") & "' order by provincia asc")
        do while not rs.eof
            response.Write("<option value=""" & rs("id") & """")
            if cInt(provincia) = cInt(rs("id")) then response.Write(" selected")
            response.Write(">" & rs("provincia") & "</option>")
            rs.movenext
        loop
        cn.close
        %>
      </select>
      </span>
      <input type="hidden" name="provincia" id="provincia" value="<%= provincia %>" alt="obbligatorio"></td>
  </tr>
  <tr>
    <td><label for="tel">telefono <strong>FISSO</strong>*<br/>
        0 se non ha fisso</label></td>
    <td><input type="text" id="tel" name="tel" value="<%= tel %>" style="width:200px" alt="obbligatorio" maxlength="50" onKeyPress="return controllonumtel09()">
      <%
      if len(tel) then

          if (tel = "0" OR tel = "00" OR tel = "000" OR tel = "0000" OR tel = "00000" OR tel = "000000" OR tel = "0000000" OR tel = "00000000" OR tel = "000000000" OR tel = "0000000000" )then
      %>
      <span style=" font-size:15px; font-weight:bold; color:#3C0;">OK - No Fisso</span>
      <%
      elseif len(tel) > 4 then
          set rsCount_Attivita_tel_Fisso = cn3.execute("select count(ID) from omaggi where tel LIKE '%" & tel & "%' OR cell LIKE '%" & tel & "%' OR cell2 LIKE '%" & tel & "%' OR cell3 LIKE '%" & tel & "%'")

          if rsCount_Attivita_tel_Fisso(0) = 1 then
      %>
      <span style=" font-size:15px; font-weight:bold; color:#3C0;">OK</span>
      <% elseif rsCount_Attivita_tel_Fisso(0) > 1 then %>
      <br/>
      <a href="https://mrecho.booksprintedizioni.it/mrecho/omaggi/?tel=<%= tel %>" target="_blank" class="lampeggia">!!! NUMERO GI&Agrave; PRESENTE IN ATTIVIT&Agrave; !!!</a><br/>
      <br/>
      <% end if %>
      <% else %>
      <span style="font-size: 15px; font-weight: bold; color: #F60;">CAMPO ERRATO!</span>
      <% end if %>
      <% end if %>
  </tr>
  <tr>
    <td><label for="email">email * </label></td>
    <td><input type="text" id="email" name="email" value="<%= email %>" style="width:340px" alt="obbligatorio" maxlength="255">
      <%
      emailmultiple = replace(email," - "," ; ")

      emailmultiple = replace(emailmultiple," ; ",";")

      emailDest = split(emailmultiple,";")

      for x = 0 to uBound(emailDest)

          if len (email) then

              set rsCount_Attivita = cn3.execute("select count(ID) from omaggi where email LIKE '%" & emailDest(x) & "%'")

              if rsCount_Attivita(0) = 1 then
      %>
      <span style=" font-size:15px; font-weight:bold; color:#3C0; float:left;">OK &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
      <% else %>
      <% if rsCount_Attivita(0) > 1 then %>
      <br/>
      <a href="https://mrecho.booksprintedizioni.it/mrecho/omaggi/?q=<%= emailDest(x) %>" target="_blank" class="lampeggia">!!! EMAIL GI&Agrave; PRESENTE IN ATTIVIT&Agrave; !!!<br/>
      >>> <%= emailDest(x) %> <<< </a><br/>
      <br/>
      <% end if %>
      <% end if %>
      <% else %>
      <span style="font-size: 15px; font-weight: bold; color: #F60;">CAMPO VUOTO!</span>
      <% end if %>
      <% next %>
  </tr>
  <tr>
    <td><label for="cell"><strong>Cellulare 1</strong></label>
      <br/></td>
    <td><input type="text" id="cell" name="cell" value="<%= cell %>"  maxlength="15" onKeyPress="return controllonumtel09()">
      <%
      if len(cell) then

          if len(cell) > 4 then
              set rsCount_Attivita_cell_1 = cn3.execute("select count(ID) from omaggi where tel LIKE '%" & cell & "%' OR cell LIKE '%" & cell & "%' OR cell2 LIKE '%" & cell & "%' OR cell3 LIKE '%" & cell & "%'")

              if rsCount_Attivita_cell_1(0) = 1 then
      %>
      <span style=" font-size:15px; font-weight:bold; color:#3C0;">OK</span>
      <% elseif rsCount_Attivita_cell_1(0) > 1 then %>
      <br/>
      <a href="https://mrecho.booksprintedizioni.it/mrecho/omaggi/?tel=<%= cell %>" target="_blank" class="lampeggia">!!! NUMERO GI&Agrave; PRESENTE IN ATTIVIT&Agrave; !!!</a><br/>
      <br/>
      <% end if %>
      <% else %>
      <span style="font-size: 15px; font-weight: bold; color: #F60;">CAMPO ERRATO!</span>
      <% end if %>
      <% end if %><br /><strong style="color: #F40000;">solo cellulari italiani, senza spazi, senza prefisso +39</strong></td>
  <tr>
  <tr>
    <td><label for="cell2"><strong>Cellulare 2</strong></label>
      <br/></td>
    <td><input type="text" id="cell2" name="cell2" value="<%= cell2 %>"  maxlength="15" onKeyPress="return controllonumtel09()">
      <%
      if len(cell2) then

          if len(cell2) > 4 then
              set rsCount_Attivita_cell_2 = cn3.execute("select count(ID) from omaggi where tel LIKE '%" & cell2 & "%' OR cell LIKE '%" & cell2 & "%' OR cell2 LIKE '%" & cell2 & "%' OR cell3 LIKE '%" & cell2 & "%'")

              if rsCount_Attivita_cell_2(0) = 1 then
      %>
      <span style=" font-size:15px; font-weight:bold; color:#3C0;">OK</span>
      <% else %>
      <% if rsCount_Attivita_cell_2(0) > 1 then %>
      <br/>
      <a href="https://mrecho.booksprintedizioni.it/mrecho/omaggi/?tel=<%= cell2 %>" target="_blank" class="lampeggia">!!! NUMERO GI&Agrave; PRESENTE IN ATTIVIT&Agrave; !!!</a><br/>
      <br/>
      <% end if %>
      <% end if %>
      <% else %>
      <span style="font-size: 15px; font-weight: bold; color: #F60;">CAMPO ERRATO!</span>
      <% end if %>
      <% end if %></td>
  <tr>
  <tr>
    <td><label for="cell3"><strong>Cellulare 3</strong></label>
      <br/></td>
    <td><input type="text" id="cell3" name="cell3" value="<%= cell3 %>" maxlength="15" onKeyPress="return controllonumtel09()">
      <%
      if len(cell3) then

          if len(cell3) > 4 then
              set rsCount_Attivita_cell_3 = cn3.execute("select count(ID) from omaggi where tel LIKE '%" & cell3 & "%' OR cell LIKE '%" & cell3 & "%' OR cell2 LIKE '%" & cell3 & "%' OR cell3 LIKE '%" & cell3 & "%'")

              if rsCount_Attivita_cell_3(0) = 1 then
      %>
      <span style=" font-size:15px; font-weight:bold; color:#3C0;">OK</span>
      <% else %>
      <% if rsCount_Attivita_cell_3(0) > 1 then %>
      <br/>
      <a href="https://mrecho.booksprintedizioni.it/mrecho/omaggi/?tel=<%= cell3 %>" target="_blank" class="lampeggia">!!! NUMERO GI&Agrave; PRESENTE IN ATTIVIT&Agrave; !!!</a><br/>
      <br/>
      <% end if %>
      <% end if %>
      <% else %>
      <span style="font-size: 15px; font-weight: bold; color: #F60;">CAMPO ERRATO!</span>
      <% end if %>
      <% end if %></td>
  <tr>
    <% cn3.close %>
  <tr>
    <td><label for="dataIscrizione">Data iscrizione</label></td>
    <td><input type="text" name="dataIscrizione" id="dataIscrizione" maxlength="10" value="<%= dataIscrizione %>" size="10" >
      <input type="button" id="dataIscrizionebtn" value=" ... ">
      (gg / mm / aaaa)</td>
  </tr>
  <tr>
    <td style="font-weight:bold;"><label for="provenienza">Provenienza</label></td>
    <td style="font-weight:bold;"><!-- verifica se è un inserimento manuale o non rilevato provenienza =0 -->
      
      <% if cInt(provenienza) = 0 or cInt(provenienza) = 5 then %>
      <select name="provenienza">
        <%
        cn.open connessioneLanding()
        set rs = cn.execute("select * from provenienza order by tipo asc")
        do while not rs.eof
            response.Write("<option value=""" & rs("tipo") & """")
            if cInt(provenienza) = rs("tipo") then response.Write("selected")
            response.Write("> " & rs("nome") & " </option>")
            rs.movenext
        loop
        cn.close
        %>
      </select>
      <%
      else
          cn.open connessioneLanding()
          set rs = cn.execute("select * from provenienza where tipo=" & provenienza)
      %>
      <input readonly="readonly" type="text"  value="<% response.write(rs("nome")) %>" style="width:320px" maxlength="255">
      <input type="hidden" id="provenienza" name="provenienza" value="<%= provenienza %>">
      <% cn.close %>
      <% end if %>
  </tr>
  <tr>
    <td><label for="dataInvioCopiaOmaggio">Data invio<br>
        copia omaggio</label></td>
    <td><input type="text" name="dataInvioCopiaOmaggio" id="dataInvioCopiaOmaggio" maxlength="10" value="<%= dataInvioCopiaOmaggio %>" size="10" >
      <input type="button" id="dataInvioCopiaOmaggiobtn" value=" ... ">
      (gg / mm / aaaa)</td>
  </tr>
  <tr>
    <td><label for="pubblicazione">Pubblicazione</label></td>
    <td><select style="width:320px;" name="pubblicazione">
        <%
        cn.open connessioneOmaggio()
        set rs = cn.execute("select * from pubblicazioni order by nome asc")
        do while not rs.eof
            response.Write("<option value=""" & rs("id") & """ ")
            if cInt(pubblicazione) = cInt(rs("id")) then response.Write(" selected")
            response.Write(">" & rs("nome") & "</option>")
            rs.movenext
        loop
        cn.close
        %>
      </select></td>
  </tr>
  <tr>
    <td><label for="predisposizione">Predisposizione</label></td>
    <td><select name="predisposizione">
        <%
        cn.open connessioneOmaggio()
        set rs = cn.execute("select * from predisposizione order by nome asc")
        do while not rs.eof
            response.Write("<option value=""" & rs("id") & """ ")
            if cInt(predisposizione) = cInt(rs("id")) then response.Write(" selected")
            response.Write(">" & rs("nome") & "</option>")
            rs.movenext
        loop
        cn.close
        %>
      </select></td>
  </tr>
   <tr>
    <td>Percorso Cartella * <br/>  <a style="color:#0080A8;" href="javascript:copycopy()">copia percorso</a></td>
    <td>
    
  <%
  if cInt(mrEchoSession("mrecho")) = 4 or cInt(mrEchoSession("mrecho")) = 6 or cInt(mrEchoSession("mrecho")) = 19 or cInt(mrEchoSession("mrecho")) = 16 or cInt(mrEchoSession("mrecho")) = 46 or cInt(mrEchoSession("mrecho")) = 18 or cInt(mrEchoSession("mrecho")) = 43 or cInt(mrEchoSession("mrecho")) = 9 or cInt(mrEchoSession("mrecho")) = 7 or cInt(mrEchoSession("mrecho")) = 2 or cInt(mrEchoSession("mrecho")) = 5 or cInt(mrEchoSession("mrecho")) = 11 or cInt(mrEchoSession("mrecho")) = 1 or cInt(mrEchoSession("mrecho")) = 47 or cInt(mrEchoSession("mrecho")) = 74 or cInt(mrEchoSession("mrecho")) = 64 or cInt(mrEchoSession("mrecho")) = 71 or cInt(mrEchoSession("mrecho")) = 75 or cInt(mrEchoSession("mrecho")) = 120 or cInt(mrEchoSession("mrecho")) = 118 or cInt(mrEchoSession("mrecho")) = 161 or cInt(mrEchoSession("mrecho")) = 117 or cInt(mrEchoSession("mrecho")) = 159 or cInt(mrEchoSession("mrecho")) = 248 or cInt(mrEchoSession("mrecho")) = 247 then
  %>
 
    <input type="text" name="percorso" id="percorso" value="<%= percorso %>" style="width:320px" alt="obbligatorio">

  <% else %>
  <input type="hidden" name="percorso" id="percorso" value="<%= percorso %>">
  <% end if %>
  
  </td>
    </tr>
  <%
  if cInt(mrEchoSession("mrecho")) = 1 or cInt(mrEchoSession("mrecho")) = 16 or cInt(mrEchoSession("mrecho")) = 19 or cInt(mrEchoSession("mrecho")) = 5 or cInt(mrEchoSession("mrecho")) = 47 or cInt(mrEchoSession("mrecho")) = 9 or cInt(mrEchoSession("mrecho")) = 12 or cInt(mrEchoSession("mrecho")) = 2 or cInt(mrEchoSession("mrecho")) = 25 or cInt(mrEchoSession("mrecho")) = 75 or cInt(mrEchoSession("mrecho")) = 159 or cInt(mrEchoSession("mrecho")) = 18 then
  %>
  <tr>
    <td>Proposte di pubblicazione</td>
    <td><input type="button" value="visualizza proposte" name="visProp" onClick="window.open('/mrecho/proposte_pubblicazione/?autore=<%= id %>','_blank')">
      <input type="button" value="nuova proposta" name="addProp" onClick="window.open('/mrecho/proposte_pubblicazione/add.asp?autore=<%= id %>','_blank')"></td>
  </tr>
  <% end if %>
  <tr>
    <td>&nbsp;</td>
    <td><%
        if isset(percorso) then
            response.Write("<a href=""file:///" & percorso & """ target=""_blank"" class=""linkTasto"">APRI PERCORSO CARTELLA</a> ")
        end if
        %></td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td><label style="color:#386F96;" for="pseudonimo"><strong>Pseudonimo</strong></label></td>
    <td><input type="text" id="pseudonimo" name="pseudonimo" value="<%= pseudonimo %>" style="width:320px" maxlength="255">
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
  <%
  if cInt(mrEchoSession("mrecho")) = 4 or cInt(mrEchoSession("mrecho")) = 18 or cInt(mrEchoSession("mrecho")) = 6 or cInt(mrEchoSession("mrecho")) = 74 or cInt(mrEchoSession("mrecho")) = 161 then
  %>
  <tr>
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3" style="background-color:#666; color:#fff; font-weight:bold">Gestione Redattrice / Grafico</td>
  </tr>
  <tr>
    <td  colspan="3"><label for="redattrice" style="margin-right:1px!important;"><strong>Redattrice:</strong></label>
      &nbsp;
      <select style="padding-right:0px!important;" name="redattrice">
        <option value="0" <% if cInt(redattrice) = 0 then response.Write("selected") %>>NON ASSEGNATO</option>
        <option value="4" <% if cInt(redattrice) = 4 then response.Write("selected") %>>Carlo</option>
        <option value="43" <% if cInt(redattrice) = 43 then response.Write("selected") %>>Angela</option>
        <option value="6" <% if cInt(redattrice) = 6 then response.Write("selected") %>>Gerarda</option>
        <option value="7" <% if cInt(redattrice) = 7 then response.Write("selected") %>>Veronica</option>
        <option value="74" <% if cInt(redattrice) = 74 then response.Write("selected") %>>Ivana</option>
        <option value="118" <% if cInt(redattrice) = 118 then response.Write("selected") %>>Gregoria</option>
        <option value="161" <% if cInt(redattrice) = 161 then response.Write("selected") %>>Maria</option>
        <option value="120" <% if cInt(redattrice) = 120 then response.Write("selected") %>>Michele + Redattrice [Note]</option>
                <option value="248" <% if cInt(redattrice) = 248 then response.Write("selected") %>>Erica</option>
      </select>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <label for="grafico" style="margin-right:1px!important;"><strong>Grafico:</strong></label>
      &nbsp;
      <select name="grafico">
        <option value="0" <% if cInt(grafico) = 0 then response.Write("selected") %>>NON ASSEGNATO</option>
        <option value="11" <% if cInt(grafico) = 11 then response.Write("selected") %>>Daniele</option>
        <option value="13" <% if cInt(grafico) = 13 then response.Write("selected") %>>Gianvito</option>
        <option value="64" <% if cInt(grafico) = 64 then response.Write("selected") %>>Maurizio</option>
        <option value="120" <% if cInt(grafico) = 120 then response.Write("selected") %>>Michele</option>
        <option value="18" <% if cInt(grafico) = 18 then response.Write("selected") %>>Nicola</option>
        <option value="46" <% if cInt(grafico) = 46 then response.Write("selected") %>>Mariano</option>
        <option value="85" <% if cInt(grafico) = 85 then response.Write("selected") %>>Federica</option>
        <option value="3" <% if cInt(grafico) = 3 then response.Write("selected") %>>Benedetto</option>
        <option value="8" <% if cInt(grafico) = 8 then response.Write("selected") %>>Carmine</option>
        <option value="117" <% if cInt(grafico) = 117 then response.Write("selected") %>>Patrizio</option>
      </select></td>
  </tr>
  <% else %>
  <tr>
    <td colspan="3"><span style="color:#F60; font-weight:bold; font-size: 14px;">Redattrice:</span> &nbsp;
      <input name="redattrice" type="hidden"  id="redattrice" style="width:270px" value="<%= redattrice %>" maxlength="300" readonly>
      <% if cInt(redattrice) = 0 then response.Write("Non Assegnato") %>
      <% if cInt(redattrice) = 4 then response.Write("Carlo") %>
      <% if cInt(redattrice) = 43 then response.Write("Angela") %>
      <% if cInt(redattrice) = 6 then response.Write("Gerarda") %>
      <% if cInt(redattrice) = 7 then response.Write("Veronica") %>
      <% if cInt(redattrice) = 74 then response.Write("Ivana") %>
      <% if cInt(redattrice) = 118 then response.Write("Gregoria") %>
      <% if cInt(redattrice) = 161 then response.Write("Maria") %>
      <% if cInt(redattrice) = 120 then response.Write("Michele + Redattrice [Note]") %>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#F60; font-weight:bold; font-size: 14px;">Grafico:</span> &nbsp;
      <input name="grafico" type="hidden"  id="grafico" style="width:270px" value="<%= grafico %>" maxlength="300" readonly>
      <% if cInt(grafico) = 0 then response.Write("Non Assegnato") %>
      <% if cInt(grafico) = 11 then response.Write("Daniele") %>
      <% if cInt(grafico) = 13 then response.Write("Gianvito") %>
      <% if cInt(grafico) = 18 then response.Write("Nicola") %>
      <% if cInt(grafico) = 64 then response.Write("Maurizio") %>
      <% if cInt(grafico) = 120 then response.Write("Michele") %>
      <% if cInt(grafico) = 46 then response.Write("Mariano") %>
      <% if cInt(grafico) = 85 then response.Write("Federica") %>
      <% if cInt(grafico) = 3 then response.Write("Benedetto") %>
      <% if cInt(grafico) = 8 then response.Write("Carmine") %>
      <% if cInt(grafico) = 117 then response.Write("Patrizio") %></td>
  </tr>
  <% end if %>
  <%
  if cInt(mrEchoSession("mrecho")) = 1 or cInt(mrEchoSession("mrecho")) = 9 or cInt(mrEchoSession("mrecho")) = 18 or cInt(mrEchoSession("mrecho")) = 3 or cInt(mrEchoSession("mrecho")) = 4 or cInt(mrEchoSession("mrecho")) = 10 or cInt(mrEchoSession("mrecho")) = 6 or cInt(mrEchoSession("mrecho")) = 7 or cInt(mrEchoSession("mrecho")) = 43 or cInt(mrEchoSession("mrecho")) = 8 or cInt(mrEchoSession("mrecho")) = 46 or cInt(mrEchoSession("mrecho")) = 120 or cInt(mrEchoSession("mrecho")) = 161 or cInt(mrEchoSession("mrecho")) = 117 then
  %>
  <br/>
  <% else %>
  <% if len(note) then %>
  <input type="hidden" name="note" value="<%= note %>">
  <% end if %>
  <% end if %>
  <tr>
    <td colspan="3" style="background-color:#666; color:#fff; font-weight:bold">Avanzamento Interno</td>
  </tr>
  <% if cInt(mrEchoSession("mrecho")) = 18 then %>
  <script type="text/javascript">
			function avanzamentoInterno(tipo){
				if(tipo == 'InternoPdfVisto'){
					attivo = prendiElementoDaId('InternoPdfVisto').checked
					if(confirm('Confermi il cambio di stato su\n InternoPdfVisto?')){
						if(prendiElementoDaId('InternoPdfVisto').checked == true){
							window.location = 'avanzamento_interno.asp?id=<%= id %>&tipo=InternoPdfVisto&InternoPdfVisto=3';
						}else{
							window.location = 'avanzamento_interno.asp?id=<%= id %>&tipo=InternoPdfVisto&InternoPdfVisto=0';
						}
					}else{
						prendiElementoDaId('InternoPdfVisto').checked = !attivo;
					}
				}
				if(tipo == 'InternoWordRevisione'){
					attivo = prendiElementoDaId('InternoWordRevisione').checked
					if(confirm('Confermi il cambio di stato su\n InternoWordRevisione?')){
						if(prendiElementoDaId('InternoWordRevisione').checked == true){
							window.location = 'avanzamento_interno.asp?id=<%= id %>&tipo=InternoWordRevisione&InternoWordRevisione=2';
						}else{
							window.location = 'avanzamento_interno.asp?id=<%= id %>&tipo=InternoWordRevisione&InternoWordRevisione=0';
						}
					}else{
						prendiElementoDaId('InternoWordRevisione').checked = !attivo;
					}
				}
			
			if(tipo == 'InternoAvviato'){
					attivo = prendiElementoDaId('InternoAvviato').checked
					if(confirm('Confermi il cambio di stato su\n InternoAvviato?')){
						if(prendiElementoDaId('InternoAvviato').checked == true){
							window.location = 'avanzamento_interno.asp?id=<%= id %>&tipo=InternoAvviato&InternoAvviato=1';
						}else{
							window.location = 'avanzamento_interno.asp?id=<%= id %>&tipo=InternoAvviato&InternoAvviato=0';
						}
					}else{
						prendiElementoDaId('InternoAvviato').checked = !attivo;
					}
				}
			
			
			
			
			
			
			}
			</script>
  <tr>
    <td colspan="2">Interno Avviato:
      <input type="checkbox" name="InternoAvviato" id="InternoAvviato" value="1" <% if cInt(InternoAvviato) = 1 then response.Write("checked") %> onClick="avanzamentoInterno('InternoAvviato')">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Inviato Word per Revisione:
      <input type="checkbox" name="InternoWordRevisione" id="InternoWordRevisione" value="2" <% if cInt(InternoWordRevisione) = 2 then response.Write("checked") %> onClick="avanzamentoInterno('InternoWordRevisione')">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Inviato Pdf per Visto:
      <input type="checkbox" name="InternoPdfVisto" id="InternoPdfVisto" value="3" <% if cInt(InternoPdfVisto) = 3 then response.Write("checked") %> onClick="avanzamentoInterno('InternoPdfVisto')"></td>
  </tr>
  <% else %>
  <tr>
    <td>Interno Avviato:</td>
    <td><%= numeroToRis(InternoAvviato) %></td>
  </tr>
  <tr>
    <td>Inviato Word per Revisione:</td>
    <td><%= numeroToRis(InternoWordRevisione) %></td>
  </tr>
  <tr>
    <td>Inviato Pdf per Visto:</td>
    <td><%= numeroToRis(InternoPdfVisto) %></td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="3" style="background-color:#666; color:#fff; font-weight:bold">Gestione Visti</td>
  </tr>
  <%
  if cInt(mrEchoSession("mrecho")) = 1 or cInt(mrEchoSession("mrecho")) = 9 or cInt(mrEchoSession("mrecho")) = 13 or cInt(mrEchoSession("mrecho")) = 18 or cInt(mrEchoSession("mrecho")) = 3 or cInt(mrEchoSession("mrecho")) = 4 or cInt(mrEchoSession("mrecho")) = 10 or cInt(mrEchoSession("mrecho")) = 6 or cInt(mrEchoSession("mrecho")) = 7 or cInt(mrEchoSession("mrecho")) = 43 or cInt(mrEchoSession("mrecho")) = 8 or cInt(mrEchoSession("mrecho")) = 46 or cInt(mrEchoSession("mrecho")) = 11 or cInt(mrEchoSession("mrecho")) = 74 or cInt(mrEchoSession("mrecho")) = 85 or cInt(mrEchoSession("mrecho")) = 64 or cInt(mrEchoSession("mrecho")) = 120 or cInt(mrEchoSession("mrecho")) = 118 or cInt(mrEchoSession("mrecho")) = 161 or cInt(mrEchoSession("mrecho")) = 117 or cInt(mrEchoSession("mrecho")) = 248 then
  %>
  <script type="text/javascript">
			function visto(tipo){
				if(tipo == 'interno'){
					attivo = prendiElementoDaId('vistoInterno').checked
					if(confirm('Confermi il cambio di stato su\nVISTO INTERNO?')){
						if(prendiElementoDaId('vistoInterno').checked == true){
							window.location = 'visto.asp?id=<%= id %>&tipo=interno&vistoInterno=1';
						}else{
							window.location = 'visto.asp?id=<%= id %>&tipo=interno&vistoInterno=0';
						}
					}else{
						prendiElementoDaId('vistoInterno').checked = !attivo;
					}
				}
				if(tipo == 'copertina'){
					attivo = prendiElementoDaId('vistoCopertina').checked
					if(confirm('Confermi il cambio di stato su\nVISTO COPERTINA?')){
						if(prendiElementoDaId('vistoCopertina').checked == true){
							window.location = 'visto.asp?id=<%= id %>&tipo=copertina&vistoCopertina=1';
						}else{
							window.location = 'visto.asp?id=<%= id %>&tipo=copertina&vistoCopertina=0';
						}
					}else{
						prendiElementoDaId('vistoCopertina').checked = !attivo;
					}
				}
			
			if(tipo == 'layoutcop'){
					attivo = prendiElementoDaId('LayoutCopertina').checked
					if(confirm('Confermi il cambio di stato su\nCONFERMA LAYOUT?')){
						if(prendiElementoDaId('LayoutCopertina').checked == true){
							window.location = 'visto.asp?id=<%= id %>&tipo=layoutcop&LayoutCopertina=1';
						}else{
							window.location = 'visto.asp?id=<%= id %>&tipo=layoutcop&LayoutCopertina=0';
						}
					}else{
						prendiElementoDaId('LayoutCopertina').checked = !attivo;
					}
				}
			
			
			
			
			
			
			}
			</script>
  <tr>
    <td colspan="2">Visto interno:
      <input type="checkbox" name="vistoInterno" id="vistoInterno" value="1" <% if cInt(vistoInterno) = 1 then response.Write("checked") %> onClick="visto('interno')">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;  Visto copertina:
      <input type="checkbox" name="vistoCopertina" id="vistoCopertina" value="1" <% if cInt(vistoCopertina) = 1 then response.Write("checked") %> onClick="visto('copertina')">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;   Visto Layout:
      <input type="checkbox" name="LayoutCopertina" id="LayoutCopertina" value="1" <% if cInt(LayoutCopertina) = 1 then response.Write("checked") %> onClick="visto('layoutcop')"></td>
  </tr>
  <tr>
    <td colspan="2" style="vertical-align:top">Note:
      <input type="text" name="note" id="note" value="<%= note %>" maxlength="300" style="width:320px">
      <input type="submit" name="salvaNota" value="salva nota"></td>
  </tr>
  <% else %>
  <tr>
    <td>Visto interno</td>
    <td><%= numeroToRis(vistoInterno) %></td>
  </tr>
  <tr>
    <td>Visto copertina</td>
    <td><%= numeroToRis(vistoCopertina) %></td>
  </tr>
  <tr>
    <td>Layout copertina</td>
    <td><%= numeroToRis(LayoutCopertina) %></td>
  </tr>
  <tr>
    <td colspan="2">Note: <%= note %></td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="2"><input type="hidden" name="id" value="<%= id %>">
      <input name="invia" type="image" id="submitButton" value="invia" src="/mrEcho/img/salva.gif" class="noborder" style="margin-top:10px;"></td>
  </tr>
</table>
<div style="display:block; width:300px; float:left; clear:both; height:15px;">&nbsp;</div>
<div style="display:block; width:300px; float:left; clear:both; height:15px;">&nbsp;</div>
<% if len(id) then %>
<form name="add" id="emailform" action="#" method="post" onsubmit="return campiObbligatori(this.name);" style="width: 100%;">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="editTable">
    <tr>
      <td colspan="3" style="background-color:#2c89ad; color:#fff; font-weight:bold">Invio Email</td>
    </tr>
    <tr>
      <td>email</td>
      <td><select style="width:320px;" name="layout" id="layout"  >
          <option value="">Seleziona...</option>
          <%
          cn.open connessioneOmaggio()
          set rs = cn.execute("select * from dbo.emailRelQ where (categoria = 1 or categoria=12) and account = " & mrEchoSession("mrecho") & " order by nome asc")
          do while not rs.eof
              response.Write("<option value=""" & rs("id") & """>" & rs("nome") & "</option>")
              rs.movenext
          loop
          cn.close
          %>
        </select></td>
    </tr>
    <tr id="titoloLibroTr">
      <td>Titolo Del Libro</td>
      <td><input type="text" name="titoloLibroEmail" id="titoloLibroEmail"  style="width:320px;"></td>
    </tr>
    <tr id="notaLibroTr">
      <td>nota</td>
      <td><input type="text" name="nota" id="nota"  style="width:320px;"></td>
    </tr>
    <tr>
      <td>allegato 1 </td>
      <td><input name="allegato1" id="allegato1" type="text" size="30" value="" maxlength="255">
        <input type="button" value="scegli" onClick="openWindow('/vbs/upload/allegato.asp?dir=allegatoEmail&campo=allegato1&nform=emailform',500,250)">
        <input type="button" value="elimina" onClick="allegato1.value = ''"></td>
    </tr>
    <tr>
      <td>allegato 2 </td>
      <td><input name="allegato2" id="allegato2" type="text" size="30" value="" maxlength="255" >
        <input type="button" value="scegli" onClick="openWindow('/vbs/upload/allegato.asp?dir=allegatoEmail&campo=allegato2&nform=emailform',500,250)">
        <input type="button" value="elimina" onClick="allegato2.value = ''"></td>
    </tr>
    <tr>
      <td>allegato 3 </td>
      <td><input name="allegato3" id="allegato3" type="text" size="30" value="" maxlength="255" >
        <input type="button" value="scegli" onClick="openWindow('/vbs/upload/allegato.asp?dir=allegatoEmail&campo=allegato3&nform=emailform',500,250)">
        <input type="button" value="elimina" onClick="allegato3.value = ''"></td>
    </tr>
    <tr>
      <td colspan="2"><script type="text/javascript" language="javascript">
					function invioEmail(){
						layout = prendiElementoDaId('layout')
						allegato1 = prendiElementoDaId('allegato1').value
						allegato2 = prendiElementoDaId('allegato2').value							
						allegato3 = prendiElementoDaId('allegato3').value
						nota = prendiElementoDaId('nota').value
						titoloLibroEmail = prendiElementoDaId('titoloLibroEmail').value
						if (confirm('confermi l\'invio del messaggio:\n'+layout.options[layout.selectedIndex].text)){
							window.location='invioemail.asp?email='+layout[layout.selectedIndex].value+'&id=<%= id %>&allegato1='+allegato1+'&allegato2='+allegato2+'&allegato3='+allegato3+'&titoloLibroEmail='+titoloLibroEmail+'&nota='+nota
						}
					}
					</script>
        <input style="background-color: #2c89ad; color: #fff; font-size: 14px; font-weight: bold; border-radius: 4px; border: none; padding: 10px 20px; margin: 10px;" type="button" value="invia email" name="invia email" onClick="invioEmail()"></td>
    </tr>
  </table>
</form>
<% end if %>
<div style="display:block; width:300px; float:left; clear:both; height:15px;">&nbsp;</div>

<!-- INVIO SMS -->
<form name="add2" id="smsform" action="#" method="post" onsubmit="return campiObbligatori(this.name);" style="width: 100%;">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="editTable">
    <tr>
      <td colspan="3" style="background-color:#3e994d; color:#fff; font-weight:bold">Invio 
        
        SMS</td>
    </tr>
    <tr>
      <td>Sms</td>
     <% if len(cell) then %> <td><select style="width:320px;" name="layoutsms" id="layoutsms"  >
          <option value="">Seleziona...</option>
          <%
          cn.open connessioneOmaggio()
          set rs = cn.execute("select * from dbo.SmsRelQ where categoria = 1 and account = " & mrEchoSession("mrecho") & " order by nome asc")
          do while not rs.eof
              response.Write("<option value=""" & rs("id") & """>" & rs("nome") & "</option>")
              rs.movenext
          loop
          cn.close
          %>
        </select></td>
    </tr>
    <tr >
      <td>Destinatario</td>
      <td><input  name="destinatarioSms" id="destinatarioSms" value="<%= cell %>" disabled></td>
    </tr>
    <tr>
 
    
      <td><script type="text/javascript" language="javascript">  
					function invioSms(){
						destinatarioSms = prendiElementoDaId('destinatarioSms').value
						layoutsms = prendiElementoDaId('layoutsms')
						LinkSms = prendiElementoDaId('LinkSms').value
						//var LinkSms = document.getElementById("LinkSms").value;
						
						if (confirm('confermi l\'invio dell\' SMS:\n'+layoutsms.options[layoutsms.selectedIndex].text+'\n\n')){
							window.location='inviosms.asp?sms='+layoutsms[layoutsms.selectedIndex].value+'&id=<%= id %>&destinatarioSms='+destinatarioSms+'&LinkSms='+LinkSms
						}
					}
										
					</script>
        <input  style="background-color: #3e994d; color: #fff; font-size: 14px; font-weight: bold; border-radius: 4px; border: none; padding: 10px 20px; margin: 10px;" type="button" value="invia sms" name="invia sms" onClick="invioSms()"></td>
      <td>&nbsp;</td>
     
         <% else %>
     
    <tr>
    <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="3"><strong style="color: #F40000;">^^^ SETTARE IL CAMPO CELLULARE NELL'APPOSITO CAMPO SOPRA ^^^<br/>senza spazi, senza prefisso internazionale (+39)</strong></td>
    </tr>
     <% end if %>
      
      
    <tr>
      <td colspan="3">&nbsp;</td>
    </tr>
  </table>
  
  <!-- GENERA Link SOCIAL -->
  <% if cInt(mrEchoSession("mrecho")) = 71 then %>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="editTable" id="shortsocial">
    <tr>
      <td colspan="2" style="background-color:#136782; color:#fff; font-weight:bold">Genera Short Link Social</td>
    </tr>
    
      <td>Link Esteso</td>
      <td><input type="text" name="longUrlSocial" id="longUrlSocial" value="<%= longUrlSocial %>" style="width:290px;">
        &nbsp;
        <input type="button" value="  riduci Link  " onClick="riduciurlsocial()"></td>
    </tr>
    <tr>
      <td>Short Link</td>
      <td><span id="LinkSmsSocial" class="short-urlsocial" style="margin-left:15px; color:#09C; font-size:11pt;"></span> <br/></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
  </table>
  <script type="text/javascript" language="javascript">
		  
		  $.bitlr({
				apiKey: '750fd7b2108f6058ddb408176db699c94a43757b',
				link: '<%= longUrlSocial %>',
				anchor: false,
				success: function(newLink) {
					
					$('.short-urlsocial').html(newLink);
				},
				error: function() {
					$('.urls').hide();
				}
				
				
			});

    function riduciurlsocial(){
				longUrlSocial = prendiElementoDaId('longUrlSocial').value		
				window.location='add.asp?id=<%= id %>&longUrlSocial='+longUrlSocial+'#shortsocial'	
			
			
						
						
						
						}

		  </script>
  <% end if %>
  
  <!-- GENERA Link -->
  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="editTable" id="genlink">
    <tr>
      <td colspan="3" style="background-color:#666; color:#fff; font-weight:bold">Short Link SMS</td>
    </tr>
    <tr>
      <td>Link Esteso</td>
      <td><input type="text" name="longUrl" id="longUrl" value="<%= longUrl %>" style="width:290px;">
        &nbsp;
        <input type="button" value="  riduci Link  " onClick="riduciurl()"></td>
    </tr>
    <tr>
      <td>>>> copia >>></td>
      <td><span class="short-url" style="margin-left:15px; color:#09C; font-size:11pt;"></span> <br/></td>
    </tr>
    <tr> 
      <script type="text/javascript" language="javascript">
		  
		  $.bitlr({
				apiKey: 'f6707a777ce97d35e8b6965c463c4e8c5092aefa',
				link: '<%= longUrl %>',
				anchor: false,
				success: function(newLink) {
					
					$('.short-url').html(newLink);
				},
				error: function() {
					$('.urls').hide();
				}
				
				
			});

    function riduciurl(){
				longUrl = prendiElementoDaId('longUrl').value		
				window.location='add.asp?id=<%= id %>&longUrl='+longUrl+'#genlink'	
			
			
						
						
						
						}

		  </script>
    <tr>
      <td>>>> incolla >>></td>
      <td><input  type="text" id="LinkSms" name="LinkSms" style="width:160px;"></td>
    </tr>
    
      <td colspan="2">&nbsp;</td>
    </tr>
  </table>
</form>
<div style="display:block; width:300px; float:left; clear:both; height:15px;">&nbsp;</div>
<div style="display:block; width:300px; float:left; clear:both; height:15px;">&nbsp;</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="editTable">
  <tr>
  
    <td colspan="3" style="<%= nessunoscontoStyle %>" >Gestione Prevendita - Nessuno Sconto</td>
  </tr>
  <%
  if cInt(mrEchoSession("mrecho")) = 1 or cInt(mrEchoSession("mrecho")) = 18 or cInt(mrEchoSession("mrecho")) = 12 or cInt(mrEchoSession("mrecho")) = 75 or cInt(mrEchoSession("mrecho")) = 141 or cInt(mrEchoSession("mrecho")) = 5 then
  %>
  <script type="text/javascript">
			function prevendita(tipo){
				if(tipo == 'pagatoprevendita'){
					attivo = prendiElementoDaId('Prevendita').checked
					if(confirm('Confermi il cambio di stato su\nPREVENDITA?')){
						if(prendiElementoDaId('Prevendita').checked == true){
							window.location = 'prevendita.asp?id=<%= id %>&tipo=pagatoprevendita&Prevendita=1';
						}else{
							window.location = 'prevendita.asp?id=<%= id %>&tipo=pagatoprevendita&Prevendita=0';
						}
					}else{
						prendiElementoDaId('Prevendita').checked = !attivo;
					}
				}
				
				
			}
			</script> 
  <script type="text/javascript">
			function nessunosconto(tipo){
				if(tipo == 'valoresconto'){
					attivo = prendiElementoDaId('NessunoSconto').checked
					if(confirm('Confermi il cambio di stato su\nNESSUNOSCONTO?')){
						if(prendiElementoDaId('NessunoSconto').checked == true){
							window.location = 'nessunosconto.asp?id=<%= id %>&tipo=valoresconto&NessunoSconto=1';
						}else{
							window.location = 'nessunosconto.asp?id=<%= id %>&tipo=valoresconto&NessunoSconto=0';
						}
					}else{
						prendiElementoDaId('NessunoSconto').checked = !attivo;
					}
				}
				
				
			}
			</script>
  <tr>
    <td colspan="2">Prevendita:
      <input type="checkbox" name="Prevendita" id="Prevendita" value="1" <% if cInt(Prevendita) = 1 then response.Write("checked") %> onClick="prevendita('pagatoprevendita')"></td>
    <td colspan="1" style="background-color:#F99!important; font-weight:bold;">Nessuno Sconto:
      <input type="checkbox" name="NessunoSconto" id="NessunoSconto" value="1" <% if cInt(NessunoSconto) = 1 then response.Write("checked") %> onClick="nessunosconto('valoresconto')"></td>
  </tr>
  <tr>
    <td colspan="3" style="vertical-align:top">Note Commerciale:<br/>
      <textarea name="noteCommerciale" rows="3" id="note" style="width:90%"><%= noteCommerciale %></textarea>
      <br/>
      <br/>
      <input type="submit" name="salvaNota" value="salva nota"></td>
  </tr>
  <% else %>
  <tr>
    <td>Prevendita</td>
    <td colspan="2"><%= numeroToRis(Prevendita) %></td>
  </tr>
  <tr>
    <td>Nessuno Sconto</td>
    <td colspan="2"><%= numeroToRis(NessunoSconto) %></td>
  </tr>
  <tr>
    <td colspan="3">Note Commerciale: <%= noteCommerciale %></td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="3">&nbsp;</td>
  </tr>
  
  
  <tr>
    <td colspan="3" style="background-color:#666; color:#fff; font-weight:bold">Gestione Documenti</td>
  </tr>
   <tr>
    <td colspan="1">Documento Identità:
      <input type="checkbox" name="DocIdentita" id="DocIdentita" value="1" <% if cInt(DocIdentita) = 1 then response.Write("checked") %> onClick="formatidigitali('DocIdentitaAttivo')"></td>
     <td colspan="2">Assunzione Responsabilità:
      <input type="checkbox" name="AssunzioneRespon" id="AssunzioneRespon" value="1" <% if cInt(AssunzioneRespon) = 1 then response.Write("checked") %> onClick="formatidigitali('AssunzioneResponAttivo')"></td>
    
  </tr>

  <tr>
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3" style="background-color:#666; color:#fff; font-weight:bold">Gestione Servizi Aggiuntivi</td>
  </tr>
 
  <script type="text/javascript">
			function formatidigitali(tipo){
				if(tipo == 'eBookAttivo'){
					attivo = prendiElementoDaId('eBook').checked
					if(confirm('Confermi il cambio di stato su\n eBook?')){
						if(prendiElementoDaId('eBook').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=eBookAttivo&eBook=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=eBookAttivo&eBook=0';
						}
					}else{
						prendiElementoDaId('eBook').checked = !attivo;
					}
				}
				
				if(tipo == 'AudioLibroAttivo'){
					attivo = prendiElementoDaId('AudioLibro').checked
					if(confirm('Confermi il cambio di stato su\nAudioLibro?')){
						if(prendiElementoDaId('AudioLibro').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=AudioLibroAttivo&AudioLibro=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=AudioLibroAttivo&AudioLibro=0';
						}
					}else{
						prendiElementoDaId('AudioLibro').checked = !attivo;
					}
				}
				
				
				
				if(tipo == 'CorrettoreBozzeAttivo'){
					attivo = prendiElementoDaId('CorrettoreBozze').checked
					if(confirm('Confermi il cambio di stato su\n Correttore Bozze?')){
						if(prendiElementoDaId('CorrettoreBozze').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=CorrettoreBozzeAttivo&CorrettoreBozze=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=CorrettoreBozzeAttivo&CorrettoreBozze=0';
						}
					}else{
						prendiElementoDaId('CorrettoreBozze').checked = !attivo;
					}
				}
				
				
				
					if(tipo == 'AssistenteRedazioneOnlineAttivo'){
					attivo = prendiElementoDaId('AssistenteRedazioneOnline').checked
					if(confirm('Confermi il cambio di stato su\n Assistente Redazione Online?')){
						if(prendiElementoDaId('AssistenteRedazioneOnline').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=AssistenteRedazioneOnlineAttivo&AssistenteRedazioneOnline=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=AssistenteRedazioneOnlineAttivo&AssistenteRedazioneOnline=0';
						}
					}else{
						prendiElementoDaId('AssistenteRedazioneOnline').checked = !attivo;
					}
				}
				
				
				
				
				
				
					if(tipo == 'CampagnaFacebookAttivo'){
					attivo = prendiElementoDaId('CampagnaFacebook').checked
					if(confirm('Confermi il cambio di stato su\n Campagna Facebook?')){
						if(prendiElementoDaId('CampagnaFacebook').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=CampagnaFacebookAttivo&CampagnaFacebook=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=CampagnaFacebookAttivo&CampagnaFacebook=0';
						}
					}else{
						prendiElementoDaId('CampagnaFacebook').checked = !attivo;
					}
				}
				
				
				
				if(tipo == 'LiveStreamingAttivo'){
					attivo = prendiElementoDaId('LiveStreaming').checked
					if(confirm('Confermi il cambio di stato su\n Live Streaming?')){
						if(prendiElementoDaId('LiveStreaming').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=LiveStreamingAttivo&LiveStreaming=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=LiveStreamingAttivo&LiveStreaming=0';
						}
					}else{
						prendiElementoDaId('LiveStreaming').checked = !attivo;
					}
				}
				
				
				if(tipo == 'DigitaleTerrestreAttivo'){
					attivo = prendiElementoDaId('DigitaleTerrestre').checked
					if(confirm('Confermi il cambio di stato su\n Digitale Terrestre?')){
						if(prendiElementoDaId('DigitaleTerrestre').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=DigitaleTerrestreAttivo&DigitaleTerrestre=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=DigitaleTerrestreAttivo&DigitaleTerrestre=0';
						}
					}else{
						prendiElementoDaId('DigitaleTerrestre').checked = !attivo;
					}
				}
				
				
				if(tipo == 'EditingAttivo'){
					attivo = prendiElementoDaId('Editing').checked
					if(confirm('Confermi il cambio di stato su\n Editing?')){
						if(prendiElementoDaId('Editing').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=EditingAttivo&Editing=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=EditingAttivo&Editing=0';
						}
					}else{
						prendiElementoDaId('Editing').checked = !attivo;
					}
				}
				
				
				
					if(tipo == 'LavorazioneEsternaAttivo'){
					attivo = prendiElementoDaId('LavorazioneEsterna').checked
					if(confirm('Confermi il cambio di stato su\n Lavorazione Esterna?')){
						if(prendiElementoDaId('LavorazioneEsterna').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=LavorazioneEsternaAttivo&LavorazioneEsterna=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=LavorazioneEsternaAttivo&LavorazioneEsterna=0';
						}
					}else{
						prendiElementoDaId('LavorazioneEsterna').checked = !attivo;
					}
				}
				
				
				
					if(tipo == 'AssunzioneResponAttivo'){
					attivo = prendiElementoDaId('AssunzioneRespon').checked
					if(confirm('Confermi il cambio di stato su\n Assunzione Responsabilità?')){
						if(prendiElementoDaId('AssunzioneRespon').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=AssunzioneResponAttivo&AssunzioneRespon=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=AssunzioneResponAttivo&AssunzioneRespon=0';
						}
					}else{
						prendiElementoDaId('AssunzioneRespon').checked = !attivo;
					}
				}
				
				
					if(tipo == 'DocIdentitaAttivo'){
					attivo = prendiElementoDaId('DocIdentita').checked
					if(confirm('Confermi il cambio di stato su\n Documento di Identità?')){
						if(prendiElementoDaId('DocIdentita').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=DocIdentitaAttivo&DocIdentita=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=DocIdentitaAttivo&DocIdentita=0';
						}
					}else{
						prendiElementoDaId('DocIdentita').checked = !attivo;
					}
				}
			
				
				
			
				
			}
			</script>
  <tr>
    <td>eBook:
      <input type="checkbox" name="eBook" id="eBook" value="1" <% if cInt(eBook) = 1 then response.Write("checked") %> onClick="formatidigitali('eBookAttivo')"></td>
     <td colspan="1">Campagna Facebook:
      <input type="checkbox" name="CampagnaFacebook" id="CampagnaFacebook" value="1" <% if cInt(CampagnaFacebook) = 1 then response.Write("checked") %> onClick="formatidigitali('CampagnaFacebookAttivo')"></td>
    <td colspan="1">Correttore Bozze:
      <input type="checkbox" name="CorrettoreBozze" id="CorrettoreBozze" value="1" <% if cInt(CorrettoreBozze) = 1 then response.Write("checked") %> onClick="formatidigitali('CorrettoreBozzeAttivo')"></td>
  </tr>
  <tr>
    <td>AudioLibro:
      <input type="checkbox" name="AudioLibro" id="AudioLibro" value="1" <% if cInt(AudioLibro) = 1 then response.Write("checked") %> onClick="formatidigitali('AudioLibroAttivo')"></td>
    <td colspan="1">Live Streaming:
      <input type="checkbox" name="LiveStreaming" id="LiveStreaming" value="1" <% if cInt(LiveStreaming) = 1 then response.Write("checked") %> onClick="formatidigitali('LiveStreamingAttivo')"></td>
    <td colspan="1">Assistente Redazione Online:
      <input type="checkbox" name="AssistenteRedazioneOnline" id="AssistenteRedazioneOnline" value="1" <% if cInt(AssistenteRedazioneOnline) = 1 then response.Write("checked") %> onClick="formatidigitali('AssistenteRedazioneOnlineAttivo')"></td>
      
  </tr>
  
  <tr>
     <td colspan="1" style="background-color: #e2ff03; font-weight: bold;">Lavorazione Esterna:
      <input type="checkbox" name="LavorazioneEsterna" id="LavorazioneEsterna" value="1" <% if cInt(LavorazioneEsterna) = 1 then response.Write("checked") %> onClick="formatidigitali('LavorazioneEsternaAttivo')"></td>
      <td colspan="1">Digitale Terrestre:
      <input type="checkbox" name="DigitaleTerrestre" id="DigitaleTerrestre" value="1" <% if cInt(DigitaleTerrestre) = 1 then response.Write("checked") %> onClick="formatidigitali('DigitaleTerrestreAttivo')"></td>
         <td colspan="1">Editing:
      <input type="checkbox" name="Editing" id="Editing" value="1" <% if cInt(Editing) = 1 then response.Write("checked") %> onClick="formatidigitali('EditingAttivo')"></td>
  </tr>
  
  
  <tr>
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3" style="background-color:#666; color:#fff; font-weight:bold">Gestione Pacchetti Ufficio Stampa</td>
  </tr>
  <%
  if cInt(mrEchoSession("mrecho")) = 1 or cInt(mrEchoSession("mrecho")) = 64 or cInt(mrEchoSession("mrecho")) = 65 or cInt(mrEchoSession("mrecho")) = 18 or cInt(mrEchoSession("mrecho")) = 71 or cInt(mrEchoSession("mrecho")) = 118 or cInt(mrEchoSession("mrecho")) = 4 or cInt(mrEchoSession("mrecho")) = 6 or cInt(mrEchoSession("mrecho")) = 7 or cInt(mrEchoSession("mrecho")) = 43 or cInt(mrEchoSession("mrecho")) = 74 or cInt(mrEchoSession("mrecho")) = 25 or cInt(mrEchoSession("mrecho")) = 100 or cInt(mrEchoSession("mrecho")) = 101 or cInt(mrEchoSession("mrecho")) = 119 or cInt(mrEchoSession("mrecho")) = 130 or cInt(mrEchoSession("mrecho")) = 131 or cInt(mrEchoSession("mrecho")) = 132 or cInt(mrEchoSession("mrecho")) = 133 or cInt(mrEchoSession("mrecho")) = 161 or cInt(mrEchoSession("mrecho")) = 141 then
  %>
  <script type="text/javascript">
			function uffstampa(tipo){
				if(tipo == 'BookTrailerAttivo'){
					attivo = prendiElementoDaId('BookTrailer').checked
					if(confirm('Confermi il cambio di stato su\n BookTrailer ?')){
						if(prendiElementoDaId('BookTrailer').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=BookTrailerAttivo&BookTrailer=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=BookTrailerAttivo&BookTrailer=0';
						}
					}else{
						prendiElementoDaId('BookTrailer').checked = !attivo;
					}
				}
				if(tipo == 'IntervistaAttivo'){
					attivo = prendiElementoDaId('Intervista').checked
					if(confirm('Confermi il cambio di stato su\nINTERVISTA AUTORE?')){
						if(prendiElementoDaId('Intervista').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=IntervistaAttivo&Intervista=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=IntervistaAttivo&Intervista=0';
						}
					}else{
						prendiElementoDaId('Intervista').checked = !attivo;
					}
				}
				
				if(tipo == 'ComunicatoStampaAttivo'){
					attivo = prendiElementoDaId('ComunicatoStampa').checked
					if(confirm('Confermi il cambio di stato su\nCOMUNICATO STAMPA?')){
						if(prendiElementoDaId('ComunicatoStampa').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=ComunicatoStampaAttivo&ComunicatoStampa=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=ComunicatoStampaAttivo&ComunicatoStampa=0';
						}
					}else{
						prendiElementoDaId('ComunicatoStampa').checked = !attivo;
					}
				}
				
				if(tipo == 'ArticoloBlogAttivo'){
					attivo = prendiElementoDaId('ArticoloBlog').checked
					if(confirm('Confermi il cambio di stato su\nARTICOLO BLOG AUTORE?')){
						if(prendiElementoDaId('ArticoloBlog').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=ArticoloBlogAttivo&ArticoloBlog=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=ArticoloBlogAttivo&ArticoloBlog=0';
						}
					}else{
						prendiElementoDaId('ArticoloBlog').checked = !attivo;
					}
				}
				
				if(tipo == 'PostFacebookAttivo'){
					attivo = prendiElementoDaId('PostFacebook').checked
					if(confirm('Confermi il cambio di stato su\nPOST FACEBOOK AUTORE?')){
						if(prendiElementoDaId('PostFacebook').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=PostFacebookAttivo&PostFacebook=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=PostFacebookAttivo&PostFacebook=0';
						}
					}else{
						prendiElementoDaId('PostFacebook').checked = !attivo;
					}
				}
				
				
					
				if(tipo == 'LibroValidoAttivo'){
					attivo = prendiElementoDaId('LibroValido').checked
					if(confirm('Confermi il cambio di stato su\nLIBRO VALIDO?')){
						if(prendiElementoDaId('LibroValido').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=LibroValidoAttivo&LibroValido=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=LibroValidoAttivo&LibroValido=0';
						}
					}else{
						prendiElementoDaId('LibroValido').checked = !attivo;
					}
				}
				
				
				
				
				
				
				
			}
			</script>
  <tr>
    <td>BookTrailer:
      <input type="checkbox" name="BookTrailer" id="BookTrailer" value="1" <% if cInt(BookTrailer) = 1 then response.Write("checked") %> onClick="uffstampa('BookTrailerAttivo')"></td>
    <td>Intervista Autore:
      <input type="checkbox" name="Intervista" id="Intervista" value="1" <% if cInt(Intervista) = 1 then response.Write("checked") %> onClick="uffstampa('IntervistaAttivo')"></td>
    <td>Comunicato Stampa:
      <input type="checkbox" name="ComunicatoStampa" id="ComunicatoStampa" value="1" <% if cInt(ComunicatoStampa) = 1 then response.Write("checked") %> onClick="uffstampa('ComunicatoStampaAttivo')"></td>
  </tr>
  <tr>
    <td>Articolo Blog:
      <input type="checkbox" name="ArticoloBlog" id="ArticoloBlog" value="1" <% if cInt(ArticoloBlog) = 1 then response.Write("checked") %> onClick="uffstampa('ArticoloBlogAttivo')"></td>
    <td>PostFacebook:
      <input type="checkbox" name="PostFacebook" id="PostFacebook" value="1" <% if cInt(PostFacebook) = 1 then response.Write("checked") %> onClick="uffstampa('PostFacebookAttivo')"></td>
    <td>Libro Valido:
      <input type="checkbox" name="LibroValido" id="LibroValido" value="1" <% if cInt(LibroValido) = 1 then response.Write("checked") %> onClick="uffstampa('LibroValidoAttivo')"></td>
  </tr>
  <% else %>
  <tr>
    <td>BookTrailer: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(BookTrailer) %></span></td>
    <td>Intervista Autore: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Intervista) %></span></td>
    <td>Comunicato Stampa: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(ComunicatoStampa) %></span></td>
  </tr>
  <tr>
    <td>Articolo Blog: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(ArticoloBlog) %></span></td>
    <td>PostFacebook: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(PostFacebook) %></span></td>
    <td>Libro Valido: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(LibroValido) %></span></td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="3">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="editTable">
  <tr>
    <td colspan="4" style="background-color:#666; color:#fff; font-weight:bold">Gestione Presenza Eventi</td>
  </tr>
  <%
  if cInt(mrEchoSession("mrecho")) = 1 or cInt(mrEchoSession("mrecho")) = 64 or cInt(mrEchoSession("mrecho")) = 65 or cInt(mrEchoSession("mrecho")) = 18 or cInt(mrEchoSession("mrecho")) = 71 or cInt(mrEchoSession("mrecho")) = 17 or cInt(mrEchoSession("mrecho")) = 25 or cInt(mrEchoSession("mrecho")) = 100 or cInt(mrEchoSession("mrecho")) = 101 or cInt(mrEchoSession("mrecho")) = 119 or cInt(mrEchoSession("mrecho")) = 130 or cInt(mrEchoSession("mrecho")) = 131 or cInt(mrEchoSession("mrecho")) = 132 or cInt(mrEchoSession("mrecho")) = 133 or cInt(mrEchoSession("mrecho")) = 141 then
  %>
  <script type="text/javascript">
			function eventi(tipo){
				if(tipo == 'Roma2015Attivo'){
					attivo = prendiElementoDaId('Roma2015').checked
					if(confirm('Confermi il cambio di stato su\n Roma 2015 ?')){
						if(prendiElementoDaId('Roma2015').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Roma2015Attivo&Roma2015=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Roma2015Attivo&Roma2015=0';
						}
					}else{
						prendiElementoDaId('Roma2015').checked = !attivo;
					}
				}
				
				
				if(tipo == 'Roma2016Attivo'){
					attivo = prendiElementoDaId('Roma2016').checked
					if(confirm('Confermi il cambio di stato su\n Roma 2016 ?')){
						if(prendiElementoDaId('Roma2016').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Roma2016Attivo&Roma2016=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Roma2016Attivo&Roma2016=0';
						}
					}else{
						prendiElementoDaId('Roma2016').checked = !attivo;
					}
				}
				
				if(tipo == 'Roma2017Attivo'){
					attivo = prendiElementoDaId('Roma2017').checked
					if(confirm('Confermi il cambio di stato su\n Roma 2017 ?')){
						if(prendiElementoDaId('Roma2017').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Roma2017Attivo&Roma2017=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Roma2017Attivo&Roma2017=0';
						}
					}else{
						prendiElementoDaId('Roma2017').checked = !attivo;
					}
				}
				
				
				if(tipo == 'Torino2015Attivo'){
					attivo = prendiElementoDaId('Torino2015').checked
					if(confirm('Confermi il cambio di stato su\n Torino 2015 ?')){
						if(prendiElementoDaId('Torino2015').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Torino2015Attivo&Torino2015=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Torino2015Attivo&Torino2015=0';
						}
					}else{
						prendiElementoDaId('Torino2015').checked = !attivo;
					}
				}
				
				
				if(tipo == 'Torino2016Attivo'){
					attivo = prendiElementoDaId('Torino2016').checked
					if(confirm('Confermi il cambio di stato su\n Torino 2016 ?')){
						if(prendiElementoDaId('Torino2016').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Torino2016Attivo&Torino2016=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Torino2016Attivo&Torino2016=0';
						}
					}else{
						prendiElementoDaId('Torino2016').checked = !attivo;
					}
				}
				
				if(tipo == 'Torino2017Attivo'){
					attivo = prendiElementoDaId('Torino2017').checked
					if(confirm('Confermi il cambio di stato su\n Torino 2017 ?')){
						if(prendiElementoDaId('Torino2017').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Torino2017Attivo&Torino2017=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Torino2017Attivo&Torino2017=0';
						}
					}else{
						prendiElementoDaId('Torino2017').checked = !attivo;
					}
				}
				
				if(tipo == 'Torino2018Attivo'){
					attivo = prendiElementoDaId('Torino2018').checked
					if(confirm('Confermi il cambio di stato su\n Torino 2018 ?')){
						if(prendiElementoDaId('Torino2018').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Torino2018Attivo&Torino2018=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Torino2018Attivo&Torino2018=0';
						}
					}else{
						prendiElementoDaId('Torino2018').checked = !attivo;
					}
				}
				
				
					if(tipo == 'Torino2019Attivo'){
					attivo = prendiElementoDaId('Torino2019').checked
					if(confirm('Confermi il cambio di stato su\n Torino 2019 ?')){
						if(prendiElementoDaId('Torino2019').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Torino2019Attivo&Torino2019=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Torino2019Attivo&Torino2019=0';
						}
					}else{
						prendiElementoDaId('Torino2019').checked = !attivo;
					}
				}
				
				
				
				
				
				
				if(tipo == 'Milano2017Attivo'){
					attivo = prendiElementoDaId('Milano2017').checked
					if(confirm('Confermi il cambio di stato su\n Milano 2017 ?')){
						if(prendiElementoDaId('Milano2017').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Milano2017Attivo&Milano2017=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Milano2017Attivo&Milano2017=0';
						}
					}else{
						prendiElementoDaId('Milano2017').checked = !attivo;
					}
				}
				
								if(tipo == 'Sanremo2013Attivo'){
					attivo = prendiElementoDaId('Sanremo2013').checked
					if(confirm('Confermi il cambio di stato su\n Sanremo 2013 ?')){
						if(prendiElementoDaId('Sanremo2013').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2013Attivo&Sanremo2013=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2013Attivo&Sanremo2013=0';
						}
					}else{
						prendiElementoDaId('Sanremo2013').checked = !attivo;
					}
				}
                
                
                if(tipo == 'Sanremo2014Attivo'){
					attivo = prendiElementoDaId('Sanremo2014').checked
					if(confirm('Confermi il cambio di stato su\n Sanremo 2014 ?')){
						if(prendiElementoDaId('Sanremo2014').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2014Attivo&Sanremo2014=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2014Attivo&Sanremo2014=0';
						}
					}else{
						prendiElementoDaId('Sanremo2014').checked = !attivo;
					}
				}
                
                if(tipo == 'Sanremo2015Attivo'){
					attivo = prendiElementoDaId('Sanremo2015').checked
					if(confirm('Confermi il cambio di stato su\n Sanremo 2015 ?')){
						if(prendiElementoDaId('Sanremo2015').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2015Attivo&Sanremo2015=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2015Attivo&Sanremo2015=0';
						}
					}else{
						prendiElementoDaId('Sanremo2015').checked = !attivo;
					}
				}
                
                 if(tipo == 'Sanremo2016Attivo'){
					attivo = prendiElementoDaId('Sanremo2016').checked
					if(confirm('Confermi il cambio di stato su\n Sanremo 2016 ?')){
						if(prendiElementoDaId('Sanremo2016').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2016Attivo&Sanremo2016=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2016Attivo&Sanremo2016=0';
						}
					}else{
						prendiElementoDaId('Sanremo2016').checked = !attivo;
					}
				}
				
				
				
				  if(tipo == 'Sanremo2019Attivo'){
					attivo = prendiElementoDaId('Sanremo2019').checked
					if(confirm('Confermi il cambio di stato su\n Sanremo 2019 ?')){
						if(prendiElementoDaId('Sanremo2019').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2019Attivo&Sanremo2019=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2019Attivo&Sanremo2019=0';
						}
					}else{
						prendiElementoDaId('Sanremo2019').checked = !attivo;
					}
				}
				
				
				  if(tipo == 'Sanremo2020Attivo'){
					attivo = prendiElementoDaId('Sanremo2020').checked
					if(confirm('Confermi il cambio di stato su\n Sanremo 2020 ?')){
						if(prendiElementoDaId('Sanremo2020').checked == true){
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2020Attivo&Sanremo2020=1';
						}else{
							window.location = 'uffstampa.asp?id=<%= id %>&tipo=Sanremo2020Attivo&Sanremo2020=0';
						}
					}else{
						prendiElementoDaId('Sanremo2020').checked = !attivo;
					}
				}
				
				
				
							
				
				
				
			}
			</script>
  <tr>
    <td>Sanremo 2013:
      <input type="checkbox" name="Sanremo2013" id="Sanremo2013" value="1" <% if cInt(Sanremo2013) = 1 then response.Write("checked") %> onClick="eventi('Sanremo2013Attivo')"></td>
    <td>Roma 2015:
      <input type="checkbox" name="Roma2015" id="Roma2015" value="1" <% if cInt(Roma2015) = 1 then response.Write("checked") %> onClick="eventi('Roma2015Attivo')"></td>
    <td>Torino 2015:
      <input type="checkbox" name="Torino2015" id="Torino2015" value="1" <% if cInt(Torino2015) = 1 then response.Write("checked") %> onClick="eventi('Torino2015Attivo')"></td>
    <td>Milano 2017:
      <input type="checkbox" name="Milano2017" id="Milano2017" value="1" <% if cInt(Milano2017) = 1 then response.Write("checked") %> onClick="eventi('Milano2017Attivo')"></td>
  </tr>
  <tr>
    <td>Sanremo 2014:
      <input type="checkbox" name="Sanremo2014" id="Sanremo2014" value="1" <% if cInt(Sanremo2014) = 1 then response.Write("checked") %> onClick="eventi('Sanremo2014Attivo')"></td>
    <td>Roma 2016:
      <input type="checkbox" name="Roma2016" id="Roma2016" value="1" <% if cInt(Roma2016) = 1 then response.Write("checked") %> onClick="eventi('Roma2016Attivo')"></td>
    <td>Torino 2016:
      <input type="checkbox" name="Torino2016" id="Torino2016" value="1" <% if cInt(Torino2016) = 1 then response.Write("checked") %> onClick="eventi('Torino2016Attivo')"></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Sanremo 2015:
      <input type="checkbox" name="Sanremo2015" id="Sanremo2015" value="1" <% if cInt(Sanremo2015) = 1 then response.Write("checked") %> onClick="eventi('Sanremo2015Attivo')"></td>
    <td>Roma 2017:
      <input type="checkbox" name="Roma2017" id="Roma2017" value="1" <% if cInt(Roma2017) = 1 then response.Write("checked") %> onClick="eventi('Roma2017Attivo')"></td>
    <td>Torino 2017:
      <input type="checkbox" name="Torino2017" id="Torino2017" value="1" <% if cInt(Torino2017) = 1 then response.Write("checked") %> onClick="eventi('Torino2017Attivo')"></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Sanremo 2016:
      <input type="checkbox" name="Sanremo2016" id="Sanremo2016" value="1" <% if cInt(Sanremo2016) = 1 then response.Write("checked") %> onClick="eventi('Sanremo2016Attivo')"></td>
    <td>&nbsp;</td>
    <td>Torino 2018:
      <input type="checkbox" name="Torino2018" id="Torino2018" value="1" <% if cInt(Torino2018) = 1 then response.Write("checked") %> onClick="eventi('Torino2018Attivo')"></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Sanremo 2019:
      <input type="checkbox" name="Sanremo2019" id="Sanremo2019" value="1" <% if cInt(Sanremo2019) = 1 then response.Write("checked") %> onClick="eventi('Sanremo2019Attivo')"></td>
    <td>&nbsp;</td>
    <td>Torino 2019:
      <input type="checkbox" name="Torino2019" id="Torino2019" value="1" <% if cInt(Torino2019) = 1 then response.Write("checked") %> onClick="eventi('Torino2019Attivo')"></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Sanremo 2020:
      <input type="checkbox" name="Sanremo2020" id="Sanremo2020" value="1" <% if cInt(Sanremo2020) = 1 then response.Write("checked") %> onClick="eventi('Sanremo2020Attivo')"></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <% else %>
  <tr>
    <td>Sanremo 2013: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Sanremo2013) %></span></td>
    <td>Roma 2015: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Roma2015) %></span></td>
    <td>Torino 2015: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Torino2015) %></span></td>
    <td>Milano 2017: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Milano2017) %></span></td>
  </tr>
  <tr>
    <td>Sanremo 2014: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Sanremo2014) %></span></td>
    <td>Roma 2016: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Roma2016) %></span></td>
    <td>Torino 2016: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Torino2016) %></span></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Sanremo 2015: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Sanremo2015) %></span></td>
    <td>Roma 2017: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Roma2017) %></span></td>
    <td>Torino 2017: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Torino2017) %></span></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Sanremo 2016: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Sanremo2016) %></span></td>
    <td>&nbsp;</td>
    <td>Torino 2018: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Torino2018) %></span></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Sanremo 2019: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Sanremo2019) %></span></td>
    <td>&nbsp;</td>
    <td>Torino 2019: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Torino2019) %></span></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>Sanremo 2020: &nbsp;&nbsp; <span style="font-weight:bold;"><%= numeroToRis(Sanremo2020) %></span></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="4">&nbsp;</td>
  </tr>
</table>
</form>
</div>

<!--INIZIO ATTIVITA-->
<div style="position:relative; top:0px; left:0px; float:left; width:500px;">
  <%
  if len(id) then
      set cn1 = Server.CreateObject("ADODB.Connection")
      idAttivita = CleanQueryString("attivita")
      style = "none"
      if isset(idAttivita) then
          cn.open connessioneOmaggio()
          set rs = cn.execute("select * from attivita where id = " & idAttivita)
          attivitaData = formatDateTime(rs("data"),2)

          attivitaNota = "..."
          if isset(rs("nota")) then attivitaNota = rs("nota")

          destinatario = rs("destinatario")
          'style = "block"
          cn.close
  %>
  <div id="linkx" style="padding:10px; background-color:#ccc;">Aggiorna attività</div>
  <span id="add1" style="">
  <form name="add1F" id="add1F" action="addAttivitaConf.asp" method="post" onsubmit="return campiObbligatori(this.name);">
    <table width="500" border="0" cellspacing="0" cellpadding="0" class="editTable">
      <tr>
        <td>data</td>
        <td><%= attivitaData %></td>
      </tr>
      <tr>
        <td>Nota</td>
        <td><%= attivitaNota %></td>
      </tr>
      <tr>
        <td colspan="2" style="background-color:#ccc">Risposta</td>
      </tr>
      <tr>
        <td>data</td>
        <td><%= now() %>
          <input type="hidden" name="data" id="dataAttivita" alt="obbligatorio" value="<%= now() %>"></td>
      </tr>
      <tr>
        <td>nota</td>
        <td><textarea spellcheck="true" id="nota1" name="nota"  class="editorPRO" rows="15" cols="80" style="width:100%"></textarea></td>
      </tr>
      <tr>
        <td colspan="2"><input type="hidden" name="svolto" value="1">
          <input type="hidden" name="idAttivita" value="<%= idAttivita %>" />
          <input type="hidden" name="id" value="<%= id %>" />
          <input name="invia" type="image" id="submitButton" value="invia" src="/mrEcho/img/salva.gif" class="noborder"></td>
      </tr>
    </table>
    <br>
    &nbsp;<br>
  </form>
  </span>
  <% else %>
  <div id="linkx" style="padding:10px; background-color:#ccc;"> <a href="javascript:visualizza()" class="linkNormale">[+] Attività svolte</a><a href="javascript:visualizza1()" class="linkNormale">[+] Nuova attività</a> </div>
  <span id="add1" style="display:none">
  <form name="add1F" id="add1F" action="addAttivitaConf.asp" method="post" onsubmit="return campiObbligatori(this.name);">
    <table width="500" border="0" cellspacing="0" cellpadding="0" class="editTable">
      <tr>
        <td>data</td>
        <td><input name="data" id="dataAttivita" alt="obbligatorio" value="<%= now() %>">
          <input type="button" id="dataAttivitabtn" value=" ... "></td>
      </tr>
      <tr>
        <td>nota</td>
        <td><textarea id="nota1" name="nota" class="editorPRO" rows="15" cols="80"  style="width:100%"></textarea></td>
      </tr>
      <tr>
        <td colspan="2"><input type="hidden" name="svolto" value="1">
          <input type="hidden" name="id" value="<%= id %>" />
          <input name="invia" type="image" id="submitButton" value="invia" src="/mrEcho/img/salva.gif" class="noborder"></td>
      </tr>
    </table>
  </form>
  <br>
  &nbsp;<br>
  </span> <span id="add2" style="display:<%= style %>">
  <form name="add2F" id="add2F" action="addAttivitaConf.asp" method="post" onsubmit="return campiObbligatori(this.name);">
    <table width="500" border="0" cellspacing="0" cellpadding="0" class="editTable">
      <tr>
        <td>data</td>
        <td><input style="width: 120px;" name="data" id="dataAttivita1"  value="<%= attivitaData %>">
          <input type="button" id="dataAttivita1btn" value=" ... ">
          &nbsp;&nbsp;
          <input type="button" value="+1" name="datafx" onClick="javascript:setDatafx(1)">
          &nbsp;
          <input type="button" value="+2" name="datafx" onClick="javascript:setDatafx(2)">
          &nbsp;
          <input type="button" value="+3" name="datafx" onClick="javascript:setDatafx(3)">
          &nbsp;
          <input type="button" value="+4" name="datafx" onClick="javascript:setDatafx(4)">
          &nbsp;
          <input type="button" value="+5" name="datafx" onClick="javascript:setDatafx(5)">
          &nbsp;
          <input type="button" value="+15" name="datafx" onClick="javascript:setDatafx(15)">
          &nbsp;
          <input type="button" value="+30" name="datafx" onClick="javascript:setDatafx(30)"></td>
      </tr>
      <tr>
        <td>layout</td>
        <td><select name="layout" id="layout"  >
            <option value="">Seleziona un template se necessario...</option>
            <%
            cn.open connessioneOmaggio()
            set rs = cn.execute("select * from dbo.emailRelQ where categoria = 5 and account = " & mrEchoSession("mrecho") & " order by nome asc")

            do while not rs.eof
                response.Write("<option value=""" & rs("id") & """>" & rs("nome") & "</option>")
                rs.movenext
            loop
            cn.close
            %>
          </select></td>
      </tr>
      <tr>
        <td>nota</td>
        <td><textarea spellcheck="true" id="nota2" name="nota" class="editorPRO" rows="15" cols="80"  style="width:100%" ><%= attivitaNota %></textarea></td>
      </tr>
      <tr>
        <td>Destinatario</td>
        <td><select name="destinatario">
            <option value="">Nessuno</option>
            <%
            destinatario = CleanQueryString("destinatario")
            cn1.open connessioneMrEcho()
            'set rs = cn1.execute("select * from userList order by [user] asc")
            set rs = cn1.execute("select * from userList where gruppo = 1 order by [user] asc")
            response.Write("<optgroup label=""Interni"">")
            do while not rs.eof
                response.Write("<option value=""" & rs("id") & """")
                if isset(destinatario) then if cInt(destinatario) = cInt(rs("id")) then response.Write(" selected")
                response.Write(">" & rs("nome") & " " & rs("cognome") & "</option>")
                rs.movenext
            loop
            response.Write("</optgroup>")

            set rs = cn1.execute("select * from userList where gruppo = 2 and CorrettoreBozze = 1 order by [user] asc")
            response.Write("<optgroup label=""Correttore Bozze"">")
            do while not rs.eof
                response.Write("<option value=""" & rs("id") & """")
                if isset(destinatario) then if cInt(destinatario) = cInt(rs("id")) then response.Write(" selected")
                response.Write(">" & rs("nome") & " " & rs("cognome") & "</option>")
                rs.movenext
            loop
            response.Write("</optgroup>")

            set rs = cn1.execute("select * from userList where gruppo = 2 and Recensionista = 1 order by [user] asc")
            response.Write("<optgroup label=""Recensionista"">")
            do while not rs.eof
                response.Write("<option value=""" & rs("id") & """")
                if isset(destinatario) then if cInt(destinatario) = cInt(rs("id")) then response.Write(" selected")
                response.Write(">" & rs("nome") & " " & rs("cognome") & "</option>")
                rs.movenext
            loop
            response.Write("</optgroup>")

            set rs = cn1.execute("select * from userList where gruppo = 2 and RedattriceOnline = 1 order by [user] asc")
            response.Write("<optgroup label=""Redattrice Online"">")
            do while not rs.eof
                response.Write("<option value=""" & rs("id") & """")
                if isset(destinatario) then if cInt(destinatario) = cInt(rs("id")) then response.Write(" selected")
                response.Write(">" & rs("nome") & " " & rs("cognome") & "</option>")
                rs.movenext
            loop
            response.Write("</optgroup>")

            set rs = cn1.execute("select * from userList where gruppo = 2 and ComunicatiStampa = 1 order by [user] asc")
            response.Write("<optgroup label=""Comunicati Stampa"">")
            do while not rs.eof
                response.Write("<option value=""" & rs("id") & """")
                if isset(destinatario) then if cInt(destinatario) = cInt(rs("id")) then response.Write(" selected")
                response.Write(">" & rs("nome") & " " & rs("cognome") & "</option>")
                rs.movenext
            loop
            response.Write("</optgroup>")

            set rs = cn1.execute("select * from userList where gruppo = 2 and Speaker = 1 order by [user] asc")
            response.Write("<optgroup style=""color:#87699;"" label=""Speaker"">")
            do while not rs.eof
                response.Write("<option value=""" & rs("id") & """")
                if isset(destinatario) then if cInt(destinatario) = cInt(rs("id")) then response.Write(" selected")
                response.Write(">" & rs("nome") & " " & rs("cognome") & "</option>")
                rs.movenext
            loop
            response.Write("</optgroup>")

            set rs = cn1.execute("select * from userList where gruppo = 2 and BookTrailer = 1 order by [user] asc")
            response.Write("<optgroup label=""BookTrailer"">")
            do while not rs.eof
                response.Write("<option value=""" & rs("id") & """")
                if isset(destinatario) then if cInt(destinatario) = cInt(rs("id")) then response.Write(" selected")
                response.Write(">" & rs("nome") & " " & rs("cognome") & "</option>")
                rs.movenext
            loop
            response.Write("</optgroup>")

            cn1.close
            %>
          </select></td>
      </tr>
      <tr>
        <td>allegato</td>
        <td><input name="allegatomrEcho" id="allegatomrEcho" type="text" size="30" value="" maxlength="255">
          <input type="button" value="scegli" onClick="openWindow('/vbs/upload/allegato.asp?dir=mrecho&campo=allegatomrEcho&nform=allegatomrEcho',500,250)">
          <input type="button" value="elimina" onClick="allegatomrEcho.value = ''"></td>
      </tr>
      <tr>
        <td>svolto</td>
        <td><select name="svolto" onChange="setData(this[this.selectedIndex].value)">
            <option value="0" <% if cInt(svolto) = 0 then response.Write("selected") %>>NO</option>
            <option value="1" <% if cInt(svolto) = 1 then response.Write("selected") %>>SI</option>
          </select></td>
      </tr>
      <tr>
        <td colspan="2"><input type="hidden" name="id" value="<%= id %>" />
          <input type="hidden" name="idAttivita" value="<%= idAttivita %>" />
          <input name="invia" type="image" id="submitButton" value="invia" src="/mrEcho/img/salva.gif" class="noborder"></td>
      </tr>
    </table>
  </form>
  <br>
  &nbsp;<br>
  </span>
  <% end if %>
  
  <!--FINE ATTIVITA--> 
  
  <span id="log">
  <div class="testata">Attività svolte</div>
  <%
  if len(mrEchoSession("mrecho")) = 0 then response.Redirect("/mrecho/login")
  id = CleanQueryString("id")
  cn1.open connessioneMrEcho()
  cn.open connessioneOmaggio()

  if len(aview) then

      if aview = "elencocompleto" then

          'visualizza elenco completo attivita svolte
          att_da_visualizzare = "*"
          nascondi_incremento_att = 1

      else
          att_da_visualizzare = "top " & Cint(aview) & " * "

      end if

  else
      att_da_visualizzare = "top 25 * "
      aview = 25

  end if

  set rs = cn.execute("select " & att_da_visualizzare & " from attivita where omaggi = " & id & " and svolto = 1 order by data desc")

  response.Write("<table border=""0"" width=""100%"" cellpadding=""2"" cellspacing=""0"" class=""elencoAttivita"">")
  response.Write("<tr><td><strong>data</strong></td><td><strong>descrizione</strong></td><td><strong>inserita_da</strong></td><td>&nbsp;</td></tr>")
  do while not rs.eof

      'Pulizia attivita senza mittente

      if IsNull(rs("user")) then
          cn.execute("update attivita set [user] = 63 where [user] is NULL and omaggi = " & id)
      end if

      assegnatodatax = formatDateTime(rs("lastMod"),2)
      datadiversa = 0

      if instr(rs("data")," ") then
          dataArr = split(rs("data")," ")
          data = formatDateTime(dataArr(0),2) & "<br>" & "<i>" & dataArr(1) & "</i>"
          if ( assegnatodatax <> formatDateTime(dataArr(0),2) ) then datadiversa = 1

      else
          data = formatDateTime(rs("data"),2)
          if ( assegnatodatax <> data ) then datadiversa = 1
      end if

      set rsUser = cn1.execute("select * from userList where id = " & rs("user"))
      if isset(rs("destinatario")) then
          set rsDest = cn1.execute("select * from userList where id = " & rs("destinatario"))
          if rsUser("user") = rsDest("user") then
              dest = ""
          else
              dest = rsDest("user")
          end if
      end if

      nota = "..."
      if isset(rs("nota")) then nota = replace(rs("nota"),vbcr,"")

      if ( datadiversa = 1) then
          response.Write("<tr id=" & rs("id") & "><td nowrap>" & data & "<span class=""assegnatodata""><br/>assegnato il:<br/>" & assegnatodatax & "</span></td><td width=""100%"">" & nota & "</td><td>" & rsUser("user") & "</td>")
      else
          response.Write("<tr id=" & rs("id") & "><td nowrap>" & data & "</td><td width=""100%"">" & nota & "</td><td>" & rsUser("user") & "</td>")
      end if

      if cInt(mrEchoSession("mrecho")) = 1 or cInt(mrEchoSession("mrecho")) = 18 then response.Write("<td><a href=""javascript:editAttivita(" & rs("id") & ")"">edit</a><br><a href=""javascript:delAttivita(" & rs("id") & ")"">del</a></td>")
      response.Write("</tr>")
      rs.movenext
  loop
  response.Write("</table>")
  cn.close
  cn1.close
  %>
  <div style=" background-color:#F9F9F9;width:500px; text-align:center; margin:0 auto; display:inline-block;">
    <%
    if nascondi_incremento_att = 0 then
        aview = Cint(aview) + 25
    %>
    <a style="color:#fff; font-weight:bold; font-size:15px;" href="./add.asp?id=<%= id %>&aview=<%= aview %>">
    <div style="padding: 12px; width: 400px; margin: 0 auto; display: inline-block; background-color:#E15D13;; margin: 10px;">Visualizza altre 25 Attività<br/>
    </div>
    </a> <a style="color:#fff; font-weight:bold; font-size:15px;" href="./add.asp?id=<%= id %>&aview=elencocompleto">
    <div style="padding: 12px; width: 400px; margin: 0 auto; display: inline-block; background-color:#C33; margin: 10px;">Visualizza Elenco Attività Completo<br/>
    </div>
    </a>
    <% else %>
    <a style="color:#fff; font-weight:bold; font-size:15px;" href="./add.asp?id=<%= id %>">
    <div style="padding: 12px; width: 400px; margin: 0 auto; display: inline-block; background-color:#399; margin: 10px;">Visualizza solo ultime 25 Attività<br/>
    </div>
    </a>
    <% end if %>
  </div>
  </span> <span id="programmate">
  <div class="testata">Attività programmate</div>
  <%
  if len(mrEchoSession("mrecho")) = 0 then response.Redirect("/mrecho/login")
  id = CleanQueryString("id")
  set cn1 = Server.CreateObject("ADODB.Connection")
  cn1.open connessioneMrEcho()
  cn.open connessioneOmaggio()
  set rs = cn.execute("select * from attivita where omaggi = " & id & " and svolto = 0 order by data desc")
  response.Write("<table border=""0"" width=""100%"" cellpadding=""2"" cellspacing=""0"" class=""elencoAttivita"">")
  response.Write("<tr><td><strong>data</strong></td><td><strong>descrizione</strong></td><td><strong>mittente</strong></td><td><strong>destinatario</strong></td><td>&nbsp;</td></tr>")
  do while not rs.eof

      'Pulizia attivita senza mittente

      if IsNull(rs("user")) then
          cn.execute("update attivita set [user] = 63 where [user] is NULL and omaggi = " & id)
      end if

      assegnatodatax = formatDateTime(rs("lastMod"),2)
      datadiversa = 0

      if instr(rs("data")," ") then
          dataArr = split(rs("data")," ")
          data = formatDateTime(dataArr(0),2) & "<br>" & "<i>" & dataArr(1) & "</i>"
          if ( assegnatodatax <> formatDateTime(dataArr(0),2) ) then datadiversa = 1
      else
          data = formatDateTime(rs("data"),2)
          if ( assegnatodatax <> data ) then datadiversa = 1
      end if
      if isset(rs("destinatario")) then
          set rsUser = cn1.execute("select * from userList where id = " & rs("destinatario"))
          set rsUser1 = cn1.execute("select * from userList where id = " & rs("user"))
          mitt = rsUser1("user")
          if rsUser("user") = mitt then mitt = "-"
      else
          set rsUser = cn1.execute("select * from userList where id = " & rs("user"))
      end if
      nota = "..."
      if isset(rs("nota")) then nota = replace(rs("nota"),vbcr,"<br>")

      if ( datadiversa = 1) then
          response.Write("<tr><td nowrap>" & data & "<br/><span class=""assegnatodata"">assegnato il:<br/>" & assegnatodatax & "</span></td><td width=""100%"" style=""vertical-align:top"">" & nota & "</td><td>" & mitt & "</td><td>" & rsUser("user") & "</td>")
      else
          response.Write("<tr><td nowrap>" & data & "</td><td width=""100%"" style=""vertical-align:top"">" & nota & "</td><td>" & mitt & "</td><td>" & rsUser("user") & "</td>")
      end if

      response.Write("<td><a href=""javascript:editAttivita(" & rs("id") & ")"">svolgi</a>")
      if cInt(mrEchoSession("mrecho")) = 1 or cInt(mrEchoSession("mrecho")) = 18 then response.Write(" <br><a href=""javascript:delAttivita(" & rs("id") & ")"">del</a>")
      response.Write("</td></tr>")
      rs.movenext
  loop
  response.Write("</table>")
  cn.close
  cn1.close
  %>
  </span>
  <div style="padding:10px; background-color:#ccc; text-align:right"> <a href="elencoAttivita.asp">vedi tutte le attività</a>&nbsp;&nbsp;&nbsp;<a href="elencoAttivita.asp?destinatario=<%= mrEchoSession("mrecho") %>">vedi le tue attività</a> </div>
  <script type="text/javascript" language="javascript">
					function Cancella_att_svolte(){
					if (confirm('confermi l\'eliminazione di tutte le attività Svolte?\n')){
							window.location='delete_att_svolte.asp?id=<%= id %>'
						}
					}
					</script> 
  <script type="text/javascript" language="javascript">
					function Cancella_att_programmate(){
					if (confirm('confermi l\'eliminazione di tutte le attività Programmate?\n')){
							window.location='delete_att_programmate.asp?id=<%= id %>'
						}
					}
					</script>
  <%
  if cInt(mrEchoSession("mrecho")) = 1 or cInt(mrEchoSession("mrecho")) = 18 or cInt(mrEchoSession("mrecho")) = 4 then
  %>
  <% if id = 4271 or id = 3640 or id = 1812 or id = 12362 or id = 58327 then %>
  <div style="padding:10px; background-color:#ccc; text-align:center">
    <input type="button" value="CANCELLA TUTTE ATTIVIT&Agrave; SVOLTE" name="CANCELLA_aTT" onClick="Cancella_att_svolte()">
  </div>
  <div style="padding:10px; background-color:#ccc; text-align:center">
    <input type="button" value="CANCELLA TUTTE ATTIVIT&Agrave; PROGRAMMATE" name="CANCELLA_aTT_PRO" onClick="Cancella_att_programmate()">
  </div>
  <% end if %>
  <% end if %>
</div>
<% end if %>
<script type="text/javascript">
		var cal = Calendar.setup({
			onSelect: function(cal) { cal.hide() },
			showTime: true,
			setLanguage:"it",
			changeAMPM: 24,
		});
		<% if len(id) then %>
		cal.manageFields("dataAttivitabtn", "dataAttivita", "%d/%m/%Y %H:%M");
		cal.manageFields("dataAttivita1btn", "dataAttivita1", "%d/%m/%Y %H:%M");
		<% end if %>
		//cal.manageFields("dataContattobtn", "dataContatto", "%d/%m/%Y");
		cal.manageFields("dataIscrizionebtn", "dataIscrizione", "%d/%m/%Y");
		cal.manageFields("dataInvioCopiaOmaggiobtn", "dataInvioCopiaOmaggio", "%d/%m/%Y");
    </script> 
<!--#include virtual="/mrecho/include/bottom1.asp"-->