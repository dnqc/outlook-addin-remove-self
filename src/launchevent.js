/*
 * RemoveSelf - Outlook Add-in
 * Rimuove automaticamente gli indirizzi del mittente dai destinatari
 * prima dell'invio, se ci sono altri destinatari nel campo A.
 * 
 * Funziona su: Outlook classico, nuovo Outlook, Outlook Web, Outlook Mobile
 * Autore: Sanitea Srl
 */

// Indirizzi da rimuovere (tutto lowercase)
var MY_ADDRESSES = [
    "d.cucciniello@sanitea.com",
    "daniele@aldero.com"
];

/**
 * Controlla se un indirizzo email è uno dei miei
 */
function isMyAddress(email) {
    if (!email) return false;
    return MY_ADDRESSES.indexOf(email.toLowerCase().trim()) !== -1;
}

/**
 * Handler principale: si attiva quando l'utente clicca Invia
 */
function onMessageSendHandler(event) {
    var item = Office.context.mailbox.item;

    // Step 1: Leggi i destinatari del campo TO
    item.to.getAsync(function (toResult) {
        if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
            // Se non riesce a leggere, lascia andare la mail
            event.completed({ allowEvent: true });
            return;
        }

        var toRecipients = toResult.value || [];

        // Step 2: Leggi i destinatari del campo CC
        item.cc.getAsync(function (ccResult) {
            if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
                event.completed({ allowEvent: true });
                return;
            }

            var ccRecipients = ccResult.value || [];

            // Step 3: Leggi i destinatari del campo BCC
            item.bcc.getAsync(function (bccResult) {
                if (bccResult.status !== Office.AsyncResultStatus.Succeeded) {
                    event.completed({ allowEvent: true });
                    return;
                }

                var bccRecipients = bccResult.value || [];

                // Conta quanti destinatari nel TO NON sono io
                var otherToCount = 0;
                for (var i = 0; i < toRecipients.length; i++) {
                    if (!isMyAddress(toRecipients[i].emailAddress)) {
                        otherToCount++;
                    }
                }

                // Se sono l'unico destinatario nel TO, non fare niente
                // (significa che sto mandando una mail a me stesso di proposito)
                if (otherToCount === 0) {
                    event.completed({ allowEvent: true });
                    return;
                }

                // Filtra i miei indirizzi da tutti i campi
                var newTo = [];
                for (var i = 0; i < toRecipients.length; i++) {
                    if (!isMyAddress(toRecipients[i].emailAddress)) {
                        newTo.push(toRecipients[i]);
                    }
                }

                var newCc = [];
                for (var i = 0; i < ccRecipients.length; i++) {
                    if (!isMyAddress(ccRecipients[i].emailAddress)) {
                        newCc.push(ccRecipients[i]);
                    }
                }

                var newBcc = [];
                for (var i = 0; i < bccRecipients.length; i++) {
                    if (!isMyAddress(bccRecipients[i].emailAddress)) {
                        newBcc.push(bccRecipients[i]);
                    }
                }

                // Controlla se è cambiato qualcosa
                var toChanged = newTo.length !== toRecipients.length;
                var ccChanged = newCc.length !== ccRecipients.length;
                var bccChanged = newBcc.length !== bccRecipients.length;

                if (!toChanged && !ccChanged && !bccChanged) {
                    // Niente da cambiare
                    event.completed({ allowEvent: true });
                    return;
                }

                // Applica le modifiche in sequenza
                setRecipientsSequential(item, newTo, newCc, newBcc, toChanged, ccChanged, bccChanged, event);
            });
        });
    });
}

/**
 * Imposta i destinatari in sequenza (TO, poi CC, poi BCC)
 * perché setAsync è asincrono
 */
function setRecipientsSequential(item, newTo, newCc, newBcc, toChanged, ccChanged, bccChanged, event) {

    // Step 1: Set TO (se cambiato)
    if (toChanged) {
        item.to.setAsync(newTo, function (result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                // Se fallisce, manda comunque la mail
                event.completed({ allowEvent: true });
                return;
            }
            // Procedi con CC
            setCcAndBcc(item, newCc, newBcc, ccChanged, bccChanged, event);
        });
    } else {
        setCcAndBcc(item, newCc, newBcc, ccChanged, bccChanged, event);
    }
}

function setCcAndBcc(item, newCc, newBcc, ccChanged, bccChanged, event) {
    // Step 2: Set CC (se cambiato)
    if (ccChanged) {
        item.cc.setAsync(newCc, function (result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                event.completed({ allowEvent: true });
                return;
            }
            // Procedi con BCC
            setBccAndComplete(item, newBcc, bccChanged, event);
        });
    } else {
        setBccAndComplete(item, newBcc, bccChanged, event);
    }
}

function setBccAndComplete(item, newBcc, bccChanged, event) {
    // Step 3: Set BCC (se cambiato)
    if (bccChanged) {
        item.bcc.setAsync(newBcc, function (result) {
            // Completa in ogni caso
            event.completed({ allowEvent: true });
        });
    } else {
        event.completed({ allowEvent: true });
    }
}

// Registra l'handler con Office
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
