var regex = /1([0-9]|-|\(|\)|\s)+/;
var string = "1 (800) 555-5555";

console.log(string.match(regex)[0]);
