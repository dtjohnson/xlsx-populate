var path = require('path');
var express = require('express');

var app = express();
var port = 8080;

console.log('Running server at http://localhost:%s/', port);

app.use('/', express.static(path.join(__dirname, 'www')));
app.use('/xlsx-populate', express.static(path.join(__dirname, '../../')));
app.listen(port);
