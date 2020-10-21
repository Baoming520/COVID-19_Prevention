var express = require('express');
var multipart = require('connect-multiparty');
var router = express.Router();

/* GET users listing. */
router.get('/', function(req, res, next) {
  res.send('respond with a resource');
});

/* POST: Use the token to exchange the password with server. */
router.post('/token', multipart(), function(req, res, next) {
  var remoteIP =  req.connection.remoteAddress;
  var reqToken = req.body.token;
  if(reqToken === '5oiR54ix5aSn54aK54yr'){
    res.send('我爱大熊猫');
  }
})

module.exports = router;
