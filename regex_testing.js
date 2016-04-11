var regex = /<span style="font-weight: bold;">/g;
var string = 'Simply <span style="font-weight: bold;">register and stay</span><span style=""> </span><span style="font-weight: bold;">6 nights </span><span style="">at +VIP Accessâ¢ hotels in 2016.</span';

string = string.replace(regex, "<b>")
string = string.replace(/<span style="">/, "");
string = string.replace(/<\/span>/, "</b>")

console.log(string);
