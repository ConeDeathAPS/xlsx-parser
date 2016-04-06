var xlsx = require("xlsx");
var fs = require("fs");
var prompt = require("prompt");

prompt.start();

var excel_doc = xlsx.readFile("./source/source.xlsx");
var spreadsheet = excel_doc.Sheets.Localization;
// console.log(spreadsheet);
var instance = new traverseColumns();
var lastColumn;
var numLocales = 0;
var numModules = 0;
var locales = [];
//The moduleDelimiters array will follow this format: [mod1startrow, mod1name, mod1endrow, mod2startrow, mod2name, mod2endrow, ...]
//The row ranges denote CONTENT locations and are inclusive. The rows containing the module names are left out of this array.
var moduleDelimiters = [];
var chars = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6, 'G': 7, 'H': 8, 'I': 9, 'J': 10, 'K': 11, 'L': 12, 'M': 13, 'N': 14, 'O': 15, 'P': 16, 'Q': 17, 'R': 18, 'S': 19, 'T': 20, 'U': 21, 'V': 22, 'W': 23, 'X': 24, 'Y': 25, 'Z': 26};
//Grab the last column with data in it from the user, since the xlsx module thinks it's some obscene value like OMF or something.
//This is also where all the functions get called
prompt.get(["last_col"], function (err, result) {
	lastColumn = result.last_col;
	//This is here because we need to wait for user input before we can execute everything else.
	//This is also the entry function for the whole parsing process.
	isolateLocales(result.last_col);
	//testing();
});

var currentCol;
var currentRow;
var pointerRow;

//This function counts the number of locales that there are based upon the last column name (it basically handles using letters as a form of counting like Excel does with column names)
//We know that there will not be more than 52 locales, so this logic is flawed if you extrapolate to all possible cases (i.e. column FJ, column DJK, etc.).
function isolateLocales(lastCol) {
	currentCol = "A";
	currentRow = "3";
	//Input validation (can't be too big or a number)
	if (lastCol.length > 2 || parseInt(lastCol)) {
		console.log("Enter the right thing! Must be a column value between A : AZ");
		return false;
	}

	//If there are two characters then we know that we have at least 26 locales
	if (lastCol.length == 2) {
		numLocales += 26;
		//Use the last char to grab any additional columns
		numLocales += (chars[lastCol.charAt(lastCol.length-1)] - 1);
	} else {
		numLocales += (chars[lastCol.charAt(lastCol.length-1)] - 1);
	}
	console.log("Found", numLocales, "locales.");
	console.log("========================================");

	var thisLocale;
	for (var i = 0; i < numLocales; i++) {
		instance.nextColumn();
		//Pull the two character codes out of all the fluff
		var regex = /[A-Z][A-Z]/g;
		var codes = spreadsheet[(currentCol+currentRow)].v.match(regex);
		if (codes.length === 1) {
			//There are a few cases of mis-formatted locale names, these are hardcoded here
			switch (codes[0]) {
				case "AR":
					thisLocale = "es_AR";
					break;

				case "NL":
					thisLocale = "nl_nL";
					break;

				case "CH":
					thisLocale = "de_CH";
					break;

				case "AT":
					thisLocale = "de_AT";
					break;

				case "TW":
					thisLocale = "zh_TW";
					break;

				case "ID":
					thisLocale = "in_ID";
					break;

				default:
					console.log("This case is not handled in locale detection!");
			}
		} else {
			//There are a few more exceptions to the locale code formatting
			switch (codes[0]+codes[1]) {
				case "FIFI":
					thisLocale = "fi_FL";
					break;

				case "VNVN":
					thisLocale = "vn_VI";
					break;

				case "KRKR":
					thisLocale = "ko_KR";
					break;

				default:
					thisLocale = codes[1].toLowerCase()+"_"+codes[0];
			}
		}
		//Each locale with have the following format:
		//{
		//name: <name>
		//langID: <id>
		//...
		//content: [
		//	{module1: [
		// 		{var1: <whatever>},
		//		{var2: <whatever>},
		//		...	
		//	]},
		//	{module2: [
		//		{var1: <whatever>},
		//		{var2: <whatever>},
		//		...
		//	]},
		//	...
		//]
		//For now, we are just initializing content and the name.
		var obj = {'name': thisLocale,
							'content': [] 
							};
		locales.push(obj);
	}
	//Check to make sure we have all the locales
	if (locales.length === numLocales) {
		//If we do, then move on to the next function
		isolateModules();
	} else {
		console.log("Error while isolating locales!");
		return 0;
	}
}
//End of locale counting function

//This function reads column A and determines modules and variable names
function isolateModules() {
	//Determine the last row. End is a middleman variable just to make this easier to read.
	var end = spreadsheet["!ref"].split(":")[1];
	var lastRow = end.match(/[0-9][0-9][0-9]/)[0];
	currentRow = "1";
	currentCol = "A";
	//handles the first module. Cannot push the very first module's end row until we find the next module's start row
	var firstMod = true;

	//Go down the column checking the strings for keywords to find the location of the modules
	while (currentRow < lastRow) {
		var cellContent = spreadsheet[(currentCol+currentRow)].v;
		if (cellContent.match(/(banner|module)/i) && !cellContent.match(/(left|right|middle|V1|V2)/i)) {
			// console.log(cellContent);
			numModules++;
			if (!firstMod) {
				moduleDelimiters.push(currentRow-1);
			}
			moduleDelimiters.push(currentRow+1);
			moduleDelimiters.push(cellContent);
			firstMod = false;
		}
		currentRow++;
		// console.log(currentRow);
	}
	console.log(numModules);
	moduleDelimiters.push(lastRow);
	// console.log(moduleDelimiters);
	console.log("Found", (moduleDelimiters.length / 3), "modules.");
	//Verify the number of modules
	if (moduleDelimiters.length / 3 === numModules) {
		//Move on to the next function
		buildLocales();
	} else {
		console.log("Error while isolating modules!");
		return 0;
	}
}

//This function uses buildModules to build out the content for each locale
function buildLocales() {
	currentCol = "B";
	// console.log(locales);
	for (var i = 0; i < 1; i++) {
		 locales[i].content = buildModules(currentCol);
		 instance.nextColumn();
	}
}

//This function parses through the spreadsheet and stores variables in each module, does one whole locale at a time then returns. Executed as called by buildLocales.
function buildModules(column) {
	//Loop through each module, we will go through each row within here creating the module for each locale
	var allMods = [];
	for (var i = 0; i < moduleDelimiters.length; i++) {
		if (moduleDelimiters[i].length > 4) {
			var modName = moduleDelimiters[i];
			var obj = {"name": modName};
			var lowerBound = moduleDelimiters[i-1];
			var upperBound = moduleDelimiters[i+1];
			for (var idx = lowerBound; idx < upperBound; idx++) {
				//Handle empty cells
				if (spreadsheet[(column+idx)]) {
					// console.log("Cell", idx , "has data");
					//Handle this case
					if (spreadsheet[(column+idx)].v === "same as US EN 1033") {
						//These look ugly, but they are just referencing and storing cell data, filling spaces with underscores, and lowercasing the letters
						obj[(spreadsheet[("A"+idx)].v).replace(/\s/g, "_").toLowerCase()] = spreadsheet[("B"+(idx.toString()))].v;
					}
					obj[(spreadsheet[("A"+idx)].v).replace(/\s/g, "_").toLowerCase()] = spreadsheet[(column+(idx.toString()))].v;
				} else {
					// console.log("Cell", idx , " does not have data");
					obj[(spreadsheet[("A"+idx)].v).replace(/\s/g, "_").toLowerCase()] = "N/A";
				}
			}
			allMods.push(obj);
		}
	}
	return allMods;
}

//This function will serve as a 'class' which must be constructed (normally as a variable called 'instance') to access .next() and .prev().
function traverseColumns() {
	//Local variables for indexing letter sequences and such
	var newCol = "";
	var charMap = {'1': 'A', '2': 'B', '3': 'C', '4': 'D', '5': 'E', '6': 'F', '7': 'G', '8': 'H', '9': 'I', '10': 'J', '11': 'K', '12': 'L', '13': 'M', '14': 'N', '15': 'O', '16': 'P', '17': 'Q', '18': 'R', '19': 'S', '20': 'T', '21': 'U', '22': 'V', '23': 'W', '24': 'X', '25': 'Y', '26': 'Z'}
	//Begin actual traversing functions. These functions can be chained (ex. instance.prev().prev().prev().next() will move the current column back 3 and forward 1)
	//Traverse forwards
	this.nextColumn = function() {
		// console.log("Moving to next column...");
		//If the current column is two characters long, the logic is slightly different
		if (currentCol.length === 2) {
			//Grab the first letter, as this isn't changing
			newCol = currentCol.charAt(0);
			//Now add the next letter, indexed via charMap which is referenced via chars
			newCol += charMap[chars[currentCol.charAt(currentCol.length - 1)] + 1];
			currentCol = newCol;
			return this;
		} else {
			//Handle the transition between 1 and 2 letter column names
			if (currentCol === "Z") {
				currentCol = "AA";
				return this;
			} else if (currentCol === lastColumn) {
				console.log("Can't move forwards from column", lastColumn, "!"); 
				return this;
			} else {
				//similar logic to above, but we aren't adding another letter, just grabbing the next letter
				newCol = charMap[chars[currentCol.charAt(0)] + 1];
				currentCol = newCol;
				return this;
			}
		}
	},
	//Traverse backwards
	this.prevColumn = function() {
		// console.log("Moving to previous column...");
		//If the current column is two characters long, the logic is slightly different
		if (currentCol.length === 2) {
			//Handle the transition between 2 and 1 letter column names
			if (currentCol === "AA") {
				currentCol = "Z";
				return this;
			}
			//Grab the first letter, as this isn't changing
			newCol = currentCol.charAt(0);
			//Now add the prev letter, indexed via charMap which is referenced via chars
			newCol += charMap[chars[currentCol.charAt(currentCol.length - 1)] - 1];
			currentCol = newCol;
			return this;
		} else {
			if (currentCol === "A") {
				console.log("Can't move backwards from column A!"); 
				return this;
			} else {
				//similar logic to above, but we aren't adding another letter, just grabbing the prev letter
				newCol = charMap[chars[currentCol.charAt(0)] - 1];
				currentCol = newCol;
				return this;
			}
		}
	}
//End of traversing functions
}

//testing and development code goes here
function testing() {

}

