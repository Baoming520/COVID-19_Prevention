var config = require('../config.json');
var express = require('express');
var fs = require('fs');
var multer = require('multer');
var path = require('path');
var redis = require('redis');
const { route } = require('./users');

var router = express.Router();
var options = { auth_pass: config.redis.password };
var cli = redis.createClient(config.redis.port, config.redis.host, options);
cli.on('connect', () => {
  console.log('[mconf_router]: Connected to Redis server.');
});

var storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, path.resolve('./data'));
  },
  filename: function (req, file, cb) {
    // cb(null, file.originalname);
    cb(null, 'mail_queue_templ.xlsx');
  }
});
var upload = multer({ storage: storage });

/* POST: Upload mail sender's configuration file */
router.post('/upload', upload.single('msconf'), function (req, res, next) {
  // var file = req.file;
  cli.hset('RC_Actions', 'UpdateMailConfig_Request', true);
  cli.hset('RC_Actions', 'UpdateMailConfig_Status', 'Pending');
  res.render('success', {
    title: 'COVID-19 Data Processing System Remote Control',
    subtitle: '- Upload MailSender App\'s config file',
    description: 'Used to update the contact information for MailSender App.',
    operation: 'UPDATE_MAIL_CONFIG',
    delay: 10
  });
});

/* GET: Download the mail sender's configuration file from website */
router.get('/', function(req, res, next){
  var data = fs.readFileSync(path.resolve('./data/mail_queue_templ.xlsx'), 'binary');
  res.setHeader('Content-Disposition', 'filename="mail_queue_templ.xlsx"');
  res.setHeader('Content-Length', data.length);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.write(data, 'binary');
  res.end('OK');
});

module.exports = router;