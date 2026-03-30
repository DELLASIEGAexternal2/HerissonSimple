/* global Office */

Office.onReady(() => {
    console.log("Add-in ready");
});

/**
 * Bouton Outlook → envoi mail simple
 */
function sendSimpleMail(event){

    Office.context.mailbox.displayNewMessageForm({

        toRecipients: ["test@email.com"],

        subject: "Test Add-in Outlook",

        htmlBody: `
        <h2>POC réussi 🚀</h2>
        <p>Email envoyé via Office.js</p>
        `
    });

    event.completed();
}

window.sendSimpleMail = sendSimpleMail;