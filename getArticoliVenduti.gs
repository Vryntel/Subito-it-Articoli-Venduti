function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Articoli Subito')
    .addItem('Leggi articoli venduti', 'getArticoliVenduti')
    .addToUi();
}


function getArticoliVenduti() {

  // Cerco email con nome articolo venduto e prezzo di vendita
  var threads = GmailApp.search('from:@messaggi.subito.it "sarà visibile sul tuo conto bancario"');
  var messages = [];

  if (threads.length == 0) {
    return;
  }

  // Se più messaggi nello stesso giorno della vendita conclusa, devo cercare in quale messaggio è presente il pagamento
  for (let i = 0; i < threads.length; i++) {
    messages = messages.concat(threads[i].getMessages());
  }

  var articoliVenduti = {};
  var idArticolo;
  var nomeArticolo;
  var prezzoVendita;
  var oggettoEmail;

  for (let i = 0; i < messages.length; i++) {
    prezzoVendita = messages[i].getPlainBody().match(/pagamento di (\d+)\s€ sarà/);

    // Se trovo il messaggio di pagamento, la mail è quella del pagamento
    if (prezzoVendita != null) {
      prezzoVendita = prezzoVendita[1];
      oggettoEmail = messages[i].getSubject();
      idArticolo = oggettoEmail.match(/\(ID:(\d+)\)/)[1];
      nomeArticolo = oggettoEmail.match(/Nuovo messaggio per (.*)\(/)[1];
      if (idArticolo != null) {
        articoliVenduti[idArticolo] = [nomeArticolo, prezzoVendita];
      }
    }
  }

  // Cerco email con transazione di vendita, info corriere e data vendita
  threads = GmailApp.search('from:noreply@subito.it subject:"Hai venduto un articolo su Subito: procedi alla spedizione"');
  messages = [];

  if (threads.length == 0) {
    return;
  }

  // Se più vendite nello stesso giorno le mail vengono unite in un thread unico, quindi per ogni thread prendo tutte le mail (messages)
  for (let i = 0; i < threads.length; i++) {
    messages = messages.concat(threads[i].getMessages());
  }


  var articoli = [];
  var urlArticolo;
  var dataVendita;
  var testoMail;
  var corriere;
  var codiceTracking;
  var urlTracking;

  idArticolo = "";
  nomeArticolo = "";
  prezzoVendita = "";

  // Prendo il codice della transazione di Subito dal nome dell'etichetta allegata, più altre info dal corpo della mail
  for (let y = 0; y < messages.length; y++) {
    urlArticolo = "https://areariservata.subito.it/transazioni/gestisci/" + messages[y].getAttachments()[0].getName().match(/etichetta-(.*)\.pdf/)[1] + "?transaction_type=fullshipping&showBackButton=true";

    dataVendita = messages[y].getDate();

    testoMail = messages[y].getPlainBody();
    codiceTracking = testoMail.match(/Il numero di vettura per questa spedizione è (.*)\./)[1];

    // Il link dell'articolo non è presente nel Plain Body
    idArticolo = messages[y].getBody().match(/https:\/\/www\.subito\.it\/vi\/(\d+)\.htm/)[1];

    if (testoMail.search("Poste Italiane") != -1) {
      corriere = "Poste Italiane";
      urlTracking = "https://www.poste.it/cerca/index.html#/risultati-spedizioni/" + codiceTracking;
    }
    else if (testoMail.search("InPost") != -1) {
      corriere = "InPost";
      urlTracking = "https://inpost.it/trova-il-tuo-pacco?number=" + codiceTracking;
    }
    else if (testoMail.search("BRT") != -1) {
      corriere = "BRT";
      urlTracking = "https://www.mybrt.it/it/mybrt/my-parcels/search?lang=it&parcelNumber=" + codiceTracking;
    }

    // Se articolo viene venduto, si crea transazione su subito, ma se compratore non ritira il pacco viene rispedito al venditore, quindi esisterà la transazione ma non la mail di effettiva vendita
    if (articoliVenduti.hasOwnProperty(idArticolo)) {
      nomeArticolo = articoliVenduti[idArticolo][0];
      prezzoVendita = articoliVenduti[idArticolo][1];
    }
    else {
      nomeArticolo = "";
      prezzoVendita = "";
    }

    articoli.push([urlArticolo, nomeArticolo, prezzoVendita, dataVendita, corriere, codiceTracking, urlTracking]);
  }

  // Ordino articoli per data di vendita
  articoli.sort(function (a, b) {
    return a[3] - b[3];
  });

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Foglio1").getRange(2, 1, articoli.length, 7).setValues(articoli);
}
