'use strict';
// Knappen i Outlook: åbner vinduet med ugen og sætter listen ind i mailen,
// når brugeren trykker "Indsæt i mail".

function aabnLedigeTider(event) {
  const url = new URL('kalender.html', location.href).href;
  let faerdig = false;
  const slut = () => { if (!faerdig) { faerdig = true; event.completed(); } };

  Office.context.ui.displayDialogAsync(url, { height: 80, width: 72, displayInIframe: false }, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) { slut(); return; }
    const vindue = res.value;
    vindue.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      let b = {};
      try { b = JSON.parse(arg.message); } catch (_) {}
      if (b.type !== 'indsaet') return;
      indsaet(b).finally(() => { vindue.close(); slut(); });
    });
    // Brugeren lukkede vinduet selv
    vindue.addEventHandler(Office.EventType.DialogEventReceived, slut);
  });
}

function indsaet(b) {
  return new Promise((ok) => {
    const krop = Office.context.mailbox.item.body;
    krop.getTypeAsync((r) => {
      const html = r.status === Office.AsyncResultStatus.Succeeded && r.value === Office.CoercionType.Html;
      krop.setSelectedDataAsync(html ? b.html : b.tekst,
        { coercionType: html ? Office.CoercionType.Html : Office.CoercionType.Text }, () => ok());
    });
  });
}

Office.onReady(() => {
  if (Office.actions) Office.actions.associate('aabnLedigeTider', aabnLedigeTider);
});
