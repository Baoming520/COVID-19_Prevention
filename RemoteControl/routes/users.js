const { v4: uuidv4 } = require('uuid');
var config = require('../config.json');
var express = require('express');
var multipart = require('connect-multiparty');
var RedisCommunicator = require('../utils/redisCommunicator');
var sendMail = require('../utils/mailer')
var util = require('util');

var router = express.Router();
var cli = RedisCommunicator.getInstance("users router", config.redis.host, config.redis.port, config.redis.password, 1).client;

/* GET users listing. */
router.get('/', function(req, res, next) {
  res.send('respond with a resource');
});

/* POST: Use the token to exchange the password with server. */
// router.post('/token', multipart(), function(req, res, next) {
//   var remoteIP =  req.connection.remoteAddress;
//   var reqToken = req.body.token;
//   if(reqToken === '5oiR54ix5aSn54aK54yr'){
//     res.send('我爱大熊猫');
//   }
// })
router.post('/getToken', multipart(), (req, res, next) => {
  var token = '口令正文'
  sendMail(req.body.mailTo, 'DEV TEST', '口令测试', token);
  res.send(token);
})

router.get('/getGuestId', (req, res, next) => {
  var reqIPAddr = req.connection.remoteAddress;
  var guestId = uuidv4();

  // Store the guest info in Redis.
  cli.hset('RC_Users', guestId, JSON.stringify({ guestId: guestId, reqIpAddr: reqIPAddr, timestamp: new Date().getTime() }));
  res.send(guestId);
});

router.post('/validate', (req, res, next) => {

});

module.exports = router;
