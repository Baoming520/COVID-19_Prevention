var config = require('../config.json');
var express = require('express');
var multer = require('multer');
var path = require('path');
var redis = require('redis');
const { route } = require('./users');

var router = express.Router();
var options = { auth_pass: config.redis.password };
var cli = redis.createClient(config.redis.port, config.redis.host, options);
cli.on('connect', () => {
  console.log('[index_router]: Connected to Redis server.');
});

/* GET home page. */
router.get('/', function (req, res, next) {
  res.render('index', { title: 'COVID-19 Data Processing System' });
});

/* GET: Send download request */
router.get('/download', function (req, res, next) {
  cli.hset('RC_Actions', 'Download_Request', true);
  cli.hset('RC_Actions', 'Download_Status', 'Pending'); // set to pending status and wait for executing by the receiver
  res.render('success', {
    title: 'COVID-19 Data Processing System Remote Control',
    subtitle: '- Download body related data',
    description: 'Download body related data of all the employees from WPS website during COVID-19.',
    operation: 'DOWNLOAD',
    delay: 3
  });
});

/* GET: Send mail sending request */
router.get('/send_mail', function (req, res, next) {
  cli.hset('RC_Actions', 'MailAll_Request', true);
  cli.hset('RC_Actions', 'MailAll_Status', 'Pending'); // set to pending status and wait for executing by the receiver
  res.render('success', {
    title: 'COVID-19 Data Processing System Remote Control',
    subtitle: '- Send mails to all contacts',
    description: 'Send mails with body related data to all contacts.',
    operation: 'MAIL_ALL',
    delay: 3
  });
});

/* GET: Send mail sending request */
router.get('/send_mail_spec', function (req, res, next) {
  cli.hset('RC_Actions', 'MailSpec_Request', true);
  cli.hset('RC_Actions', 'MailSpec_Status', 'Pending'); // set to pending status and wait for executing by the receiver
  res.render('success', {
    title: 'COVID-19 Data Processing System Remote Control',
    subtitle: '- Send mails to special contacts',
    description: 'Send mails with statistical data to the special contacts',
    operation: 'MAIL_SPEC',
    delay: 3
  });
});

/* GET: Check the results of all the operations */
router.get('/check', function (req, res, next) {
  res.render('check', {
    title: 'COVID-19 Data Processing System Remote Control',
    description: '- Check status for all the operations'
  });
});

/* GET: Read status of all the actions from db */
router.get('/actions', function (req, res, next) {
  cli.hmget('RC_Actions', [
    'Download_Request',
    'Download_Status',
    'MailAll_Request',
    'MailAll_Status',
    'MailSpec_Request',
    'MailSpec_Status',
    'UpdateMailConfig_Request',
    'UpdateMailConfig_Status'
  ], (err, res_arr) => {
    if (err) {
      console.error(err);
    }

    res.send({
      downloadRequest: res_arr[0] === 'true' ? 'Yes' : 'No',
      downloadStatus: !res_arr[1] ? 'Not Started' : res_arr[1],
      mailAllRequest: res_arr[2] === 'true' ? 'Yes' : 'No',
      mailAllStatus: !res_arr[3] ? 'Not Started' : res_arr[3],
      mailSpecRequest: res_arr[4] === 'true' ? 'Yes' : 'No',
      mailSpecStatus: !res_arr[5] ? 'Not Started' : res_arr[5],
      updateMailConfigRequest: res_arr[6] === 'true' ? 'Yes' : 'No',
      updateMailConfigStatus: !res_arr[7] ? 'Not Started' : res_arr[7]
    });
  });
});

module.exports = router;
