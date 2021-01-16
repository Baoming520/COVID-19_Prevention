var config = require('../config.json');
var mail = require('nodemailer');
var log4js = require('log4js');
var util = require('util');

// Logger
log4js.configure(config.log4js);
var logger = log4js.getLogger('default');

var transport = mail.createTransport({
  host: config.mail.host,
  secureConnection: true,
  port: config.mail.port,
  auth: {
    user: config.mail.user,
    pass: config.mail.pass
  },
  tls: {
    rejectUnauthorized: false
  }
});

var options = {
  
};

var sendMail = function(mailTo, displayName, subject, content){
  var options = {
    from: util.format('\"%s\" %s', displayName, config.mail.user),
    to: mailTo,
    subject: subject,
    text: content
  };

  transport.sendMail(options, (err, res) => {
    if(err){
      logger.error(err);
      return;
    }

    logger.info(res);
  });
};

module.exports = sendMail;