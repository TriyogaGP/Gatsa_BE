const express = require("express");
const cors = require('cors');
const morgan = require('morgan');
const query = require('./config/database');
var fs = require('fs');
var path = require('path');
let ejs = require("ejs");
let pdf = require("html-pdf");
const sendLogDB = require('../src/config/winston');
const { response } = require('./utils/response.utils');
const indexRouter = require('./routes/index');
const dotenv = require('dotenv');

const app = express();
dotenv.config();

// view engine setup
app.set('views', path.join(__dirname, 'src/views'));
app.set('view engine', 'ejs');

app.use(morgan('combined', { stream: sendLogDB.stream }));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(express.static(path.join(__dirname,'/public')));
app.use(cors({credentials:true, origin:'http://localhost:3000'}));
app.options("*", cors());

const port = Number(process.env.PORT || 3000);

indexRouter(app);

app.all('*', (req, res, next) => {
  return response(res, { kode: 404, message: 'Endpoint Not Found' }, 404);
})

// app.use(errorMiddleware);

app.listen(port, () => console.log(`Server running on port ${port} !`));

module.exports = app;