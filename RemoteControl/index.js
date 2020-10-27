var app = require('./app');
var config = require('./config.json')

app.listen(config.server.port, function(){
  console.log('Listen on 127.0.0.1:%d', config.server.port);
});