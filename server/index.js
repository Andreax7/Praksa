const express = require('express');

const app = express();
const port = 3000;
var bodyParser = require('body-parser');
var routes = require('./routes');
var excelRoute = require('./excelRoute');


//MIDDLEWARE
var cors = require('cors');
var corsOptions = {
    origin: "*",
    preflightContinue: false,
    optionsSuccessStatus: 200
  }


  app.use(cors(corsOptions));
  app.use(express.json());
  // ROUTES
  app.use("/",routes);
  app.use("/excel",excelRoute);

app.listen(port, () => {
  console.log(`API listening at http://localhost:${port}`);
});