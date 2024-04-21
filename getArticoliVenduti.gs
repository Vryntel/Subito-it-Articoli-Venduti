function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Articoli Subito')
      .addItem('Leggi articoli venduti', 'getArticoliVenduti')
      .addToUi();
}

function getArticoliVenduti() {
  var articoli = [];
  var messages = [];
  var urlArticolo;
  var dataVendita;
  var testoMail;
  var corriere;
  var codiceTracking;
  var urlTracking;

  const threads = GmailApp.search('from:noreply@subito.it subject:"Hai venduto un articolo su Subito: procedi alla spedizione"');

  if (threads.length == 0) {
    return;
  }

  // Se più vendite nello stesso giorno le mail vengono unite in un thread unico, quindi per ogni thread prendo tutte le mail (messages)
  for (let i = 0; i < threads.length; i++) {
    messages = messages.concat(threads[i].getMessages());
  }

  // Prendo il codice della transazione di Subito dal nome dell'etichetta allegata, più altre info dal corpo della mail
  for (let y = 0; y < messages.length; y++) {
    urlArticolo = "https://areariservata.subito.it/transazioni/gestisci/" + messages[y].getAttachments()[0].getName().match(/etichetta-(.*)\.pdf/)[1] + "?transaction_type=fullshipping&showBackButton=true";

    dataVendita = messages[y].getDate();
    testoMail = messages[y].getPlainBody();
    codiceTracking = testoMail.match(/Il numero di vettura per questa spedizione è (.*)\./)[1];

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

    articoli.push([urlArticolo, dataVendita, corriere, codiceTracking, urlTracking]);
  }

  // Ordino articoli per data di vendita
  articoli.sort(function (a, b) {
    return a[1] - b[1];
  });

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Foglio1").getRange(2, 1, articoli.length, 5).setValues(articoli);
}
