'use strict';

const Unimail = require('../lib');

const unimail = new Unimail();
const connectors = {};

if (process.env.GMAIL_CLIENT_ID && process.env.GMAIL_CLIENT_SECRET) {
    const gmail = new Unimail.GmailConnector({
        clientId: process.env.GMAIL_CLIENT_ID,
        clientSecret: process.env.GMAIL_CLIENT_SECRET
    });
    unimail.use(gmail);
    connectors.gmail = gmail;
}

if (process.env.OFFICE365_CLIENT_ID && process.env.OFFICE365_CLIENT_SECRET) {
    const office365 = new Unimail.Office365Connector({
        clientId: process.env.OFFICE365_CLIENT_ID,
        clientSecret: process.env.OFFICE365_CLIENT_SECRET
    });
    unimail.use(office365);
    connectors.office365 = office365;
}

if (process.env.UNIPILE_BASE_URL && process.env.UNIPILE_ACCESS_TOKEN) {
    const unipile = new Unimail.UnipileConnector({
        baseUrl: process.env.UNIPILE_BASE_URL,
        accessToken: process.env.UNIPILE_ACCESS_TOKEN
    });
    unimail.use(unipile);
    connectors.unipile = unipile;
}

module.exports = { unimail, connectors };
