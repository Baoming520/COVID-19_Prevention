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
  cli.hset('rc_operations', 'download', true);
  cli.hset('rc_operations', 'download_status', 'Pending');
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
  cli.hset('rc_operations', 'sendMail', true);
  cli.hset('rc_operations', 'sendMail_status', 'Pending');
  res.render('success', {
    title: 'COVID-19 Data Processing System Remote Control',
    subtitle: '- Send mails to all contacts',
    description: 'Send mails with body related data to all contacts.',
    operation: 'SEND_MAIL',
    delay: 3
  });
});

/* GET: Send mail sending request */
router.get('/send_mail_spec', function (req, res, next) {
  cli.hset('rc_operations', 'sendMailX', true);
  cli.hset('rc_operations', 'sendMailX_status', 'Pending');
  res.render('success', {
    title: 'COVID-19 Data Processing System Remote Control',
    subtitle: '- Send mails to special contacts',
    description: 'Send mails with statistical data to the special contacts',
    operation: 'SEND_MAIL_X',
    delay: 3
  });
});

/* GET: Check the results of all the operations */
router.get('/check', function (req, res, next) {
  cli.hmget('rc_operations', [
    'download',
    'download_status',
    'sendMail',
    'sendMail_status',
    'sendMailX',
    'sendMailX_status',
    'updateConfig',
    'updateConfig_status'
  ], (err, res_arr) => {
    if (err) {
      console.error(err);
    }

    res.render('check', {
      title: 'COVID-19 Data Processing System Remote Control',
      description: '- Check status for all the operations',
      download_request: res_arr[0] === 'true' ? 'Yes' : 'No',
      download_status: res_arr[1] === '' || res_arr[1] === null ? 'Not Started' : res_arr[1],
      sendM_request: res_arr[2] === 'true' ? 'Yes' : 'No',
      sendM_status: res_arr[3] === '' || res_arr[3] === null ? 'Not Started' : res_arr[3],
      sendMX_request: res_arr[4] === 'true' ? 'Yes' : 'No',
      sendMX_status: res_arr[5] === '' || res_arr[5] === null ? 'Not Started' : res_arr[5],
      upMConfig_request: res_arr[6] === 'true' ? 'Yes' : 'No',
      upMConfig_status: res_arr[7] === '' || res_arr[7] === null ? 'Not Started' : res_arr[7],
    });
  });
});

module.exports = router;
