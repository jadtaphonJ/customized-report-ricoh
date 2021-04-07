const express = require('express');
const app = express();

app.get('/get-report', (req, res) => {
    res.send('Hello World, jadtapyon')
})

const server = app.listen(3548, () => console.log(`Express server listening on port 3548`));

module.exports = app;