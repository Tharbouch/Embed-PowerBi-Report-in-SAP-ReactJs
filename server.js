const https = require("https");
const fs = require("fs");
const path = require('path');
const express = require("express");


const app = express();


https
    .createServer(
        {
            key: fs.readFileSync("certs/server-key.pem"),
            cert: fs.readFileSync("certs/server.pem"),
        },
        app
    )
    .listen(3000, () => {
        console.log('server is runing at port 3000')
    });


app.use(express.static(path.join(__dirname, 'build')));

app.get('/', function (req, res) {
    res.sendFile(path.join(__dirname, 'build', 'index.html'));
});
