const express = require('express');
const morgan = require('morgan');
const path = require('path');

const DEFAULT_PORT = process.env.PORT || 3000;

// initialize express.
const app = express();

// Configure morgan module to log all requests.
app.use(morgan('dev'));

// Setup app folders.
// no-store: this is an actively-edited demo — without this, a browser tab
// left open across edits can silently keep serving a stale cached
// index.html/switcher.js/etc. (express.static sends no cache headers by
// default, so browsers are free to heuristically cache anyway), making a
// real code change look like it "isn't taking effect."
app.use(express.static('public', { setHeaders: (res) => res.setHeader('Cache-Control', 'no-store') }));

// set up a route for redirect.html
app.get('/redirect', (req, res) => {
    res.setHeader('Cache-Control', 'no-store');
    res.sendFile(path.join(__dirname + '/public/redirect.html'));
});

// Set up a route for index.html
app.get('/', (req, res) => {
    res.setHeader('Cache-Control', 'no-store');
    res.sendFile(path.join(__dirname + '/index.html'));
});

app.listen(DEFAULT_PORT, () => {
    console.log(`Sample app listening on port ${DEFAULT_PORT}!`);
});

module.exports = app;
