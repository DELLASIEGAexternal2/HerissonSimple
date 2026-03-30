/* global Office */

Office.onReady(() => {});

/**
 * Envoi simple via Outlook (ouvre un mail pré-rempli)
 */
function sendSimpleMail(event) {

    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["PrimoSylvestreDELLASIEGA.external2@banque-france.fr"],

        subject: "Test Hérisson - Signalement",

        htmlBody: `
            <h2>Signalement automatique</h2>
            <p>Mail envoyé depuis le Web Add-in Hérisson</p>
        `
    });

    event.completed();
}

window.sendSimpleMail = sendSimpleMail;
