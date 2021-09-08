var winapi = require('winapi');

console.log("Last input time is %s", winapi.GetLastInputInfo() );

setTimeout(function(){
  //do not move, it wont change !
  console.log("Last input time is %s", winapi.GetLastInputInfo() );
}, 1000);