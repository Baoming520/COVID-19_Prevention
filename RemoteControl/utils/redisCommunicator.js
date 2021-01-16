const { exception, error } = require('console');
var redis = require('redis');
var util = require('util');

function RedisCommunicator(name, host, port, pass, dbIndex = 0) {
  var options = { auth_pass: pass };
  this.name = name;
  this.client = redis.createClient(port, host, options);
  this.client.select(dbIndex, () => {
    console.log(util.format('Select database: %d.', dbIndex));
  });

  this.client.on('connect', () => {
    console.log(util.format('[%s]: Connected to Redis server.', this.name));
  });
}

RedisCommunicator._names = [];
RedisCommunicator._instances = [];

RedisCommunicator.getInstance = (name, host, port, pass, dbIndex) => {
  if(!name){
    throw new Error('The name of redis communicator MUST NOT be null, undefined or empty.');
  }

  var instance = null;
  var index = RedisCommunicator._names.indexOf(name);
  if (index == -1){
    instance = new RedisCommunicator(name, host, port, pass, dbIndex);
    RedisCommunicator._instances.push(instance);
  }
  else{
    instance = RedisCommunicator._instances[index];
  }

  return instance;
};

module.exports = RedisCommunicator;