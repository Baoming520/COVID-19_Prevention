var async = require('async');
var commands = require('./commands.json');
var config = require('./config.json');
var execFile = require('child_process').execFile;
var iconv = require('iconv-lite');
var path = require('path');
var redis = require('redis');

var options = { auth_pass: config.redis.password };
var cli = redis.createClient(config.redis.port, config.redis.host, options);
cli.on('connect', () => {
  console.log('Connected to Redis server.');
});

commands.forEach(cmd => {
  var m = () => {
    async.waterfall([
      function(callback){
        cli.hget(cmd.opGroup, cmd.alias, (err, res) => {
          if(err){
            console.error(err);
            return;
          }
          callback(null, res);
        });
      },
      function(executable, callback){
        if(executable === 'true'){
          cli.hget(cmd.opGroup, cmd.alias + '_status', (err, res) => {
            if(err){
              console.error(err);
              return;
            }
            callback(null, res);
          });
        }
      },
      function(status, callback){
        if(status === 'Pending'){
          console.log('[Start]: ' + cmd.name);
          cli.hset(cmd.opGroup, cmd.alias + '_status', 'In Progress');
          execFile(path.join(cmd.env, cmd.exe), cmd.params, { encoding: 'cp936', cwd: cmd.env }, (err, res) => {
            if(err){
              console.error(err);
              return;
            }
            callback(null, res);
          });
        }
      }
    ], function(err, res){
      if (err) {
        console.error(err);
        return;
      }
      
      cli.hset(cmd.opGroup, cmd.alias, false);
      cli.hset(cmd.opGroup, cmd.alias + '_status', 'Completed');
      console.log(iconv.decode(Buffer.from(res, 'binary'), 'cp936'));
      console.log('[End]: ' + cmd.name);
    });
  };

  setInterval(m, cmd.interval);
});
