//Dependencies
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
var currentCol;
var currentRow;
var pointerRow;
//The moduleDelimiters array will follow this format: [mod1startrow, mod1name, mod1endrow, mod2startrow, mod2name, mod2endrow, ...]
//The row ranges denote CONTENT locations and are inclusive. The rows containing the module names are left out of this array.
var moduleDelimiters = [];
var chars = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6, 'G': 7, 'H': 8, 'I': 9, 'J': 10, 'K': 11, 'L': 12, 'M': 13, 'N': 14, 'O': 15, 'P': 16, 'Q': 17, 'R': 18, 'S': 19, 'T': 20, 'U': 21, 'V': 22, 'W': 23, 'X': 24, 'Y': 25, 'Z': 26};
//Grab the last column with data in it from the user, since the xlsx module thinks it's some obscene value like OMF or something.
//This is also the entry point (isolateLocales)
prompt.get(["last_col"], function (err, result) {
	lastColumn = result.last_col;
	//This is here because we need to wait for user input before we can execute everything else.
	//This is also the entry function for the whole parsing process.
	isolateLocales(result.last_col);
	//testing();
});

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
	moduleDelimiters.push(lastRow);
	// console.log(moduleDelimiters);
	console.log("Found", (moduleDelimiters.length / 3).toString(), "modules, ", numLocales.toString(), "locales...");
	console.log("============================");
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
		console.log("Building for locale", locales[i]);		
		locales[i].eyebrow_linktext = ((spreadsheet[currentCol+"23"].v === "same as CA EN 4105") ? spreadsheet["C"+"23"].v : spreadsheet[currentCol+"23"].v);
		locales[i].eyebrow_link = ((spreadsheet[currentCol+"24"].v === "same as CA EN 4105") ? spreadsheet["C"+"24"].v : spreadsheet[currentCol+"24"].v);
		locales[i].content = buildModules(currentCol);		
		instance.nextColumn();
	}
	normalizeStrings(locales);
}

//This function parses through the spreadsheet and stores variables in each module, does one whole locale at a time then returns. Executed as called by buildLocales.
function buildModules(column) {
	//Loop through each module, we will go through each row within here creating the module for each locale
	var obj = {};
	obj.locVars = {};
	for (var i = 0; i < moduleDelimiters.length; i++) {
		if (moduleDelimiters[i].length > 4) {
			// console.log(moduleDelimiters);
			var modName = moduleDelimiters[i];
			var lowerBound = moduleDelimiters[i-1];
			var upperBound = moduleDelimiters[i+1];
			//Make sure that if a new module is added, you add a case for it here!!			
			switch (modName) {
				//Construction of the RWS module
				case "Rewards Summary Module":
					//This module has it's own separate object in the loc files. Here we are initializing that as well as some static styling and img links
					obj.rewardsObj = {};
					obj.rewardsObj.style = {};
					obj.rewardsObj.style.gold = {"TierColor": "#e7b40d", "PlusImage": "http://media.expedia.com/media/content/expus/graphics/mail/general/ss_Icon_StatementSummary_GoldPlusses_71x63.jpg"};
					obj.rewardsObj.style.silver = {"TierColor": "#b1b3b6", "PlusImage": "http://media.expedia.com/media/content/expus/graphics/mail/general/ss_Icon_StatementSummary_SilverPlusses_71x63.jpg"};
					obj.rewardsObj.style.blue = {"TierColor": "#0085c9", "PlusImage": "http://media.expedia.com/media/content/expus/graphics/mail/general/ss_Icon_StatementSummary_BluePlusses_71x63.jpg"};
					//Now we can begin looping through the variables inside of each module. Each module case has an additional switch inside of it to make sure that the variable names are compatible.
					for (var idx = lowerBound; idx < upperBound; idx++) {
						//Make sure the cell has data
						if (spreadsheet[(column+idx)]) {
							//Grab the data
							var varContent = spreadsheet[(column+idx)].w;	
							//If there is data, we need to make sure we name the variable appropriately. There are lots of possibilities, so we need another switch.
							switch (spreadsheet["A"+idx].v) {
								case "Bottom bar value: gold":
									obj.rewardsObj.GVB_Max = varContent;
									break;

								case "Status Bar Copy Type":
									obj.rewardsObj.summaryType = {};
									for (var x = 1; x <= 7; x++) {
										if (varContent === "same as US EN 1033") {
											obj.rewardsObj.summaryType["subheader"+x] = spreadsheet["B"+(idx+x)].w;
										} else {
											obj.rewardsObj.summaryType["subheader"+x] = spreadsheet[column+(idx+x)].w;
										}
									}
									break;

								case '"As of" Label:':
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.asOfNightsText = spreadsheet["B"+idx].w.replace(/(\[date\]|date)/, "${AsOfDate}");
										obj.rewardsObj.asOfDollarsText = spreadsheet["B"+idx].w.replace(/(\[date\]|date)/, "${AsOfDate}");
									} else {
										obj.rewardsObj.asOfNightsText = varContent.replace(/(\[date\]|date)/, "${AsOfDate}");
										obj.rewardsObj.asOfDollarsText = varContent.replace(/(\[date\]|date)/, "${AsOfDate}");											
									}
									break;

								case "Header":
										if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryHeaderLinkText = spreadsheet["B"+idx].w;
									} else {
										obj.rewardsObj.summaryHeaderLinkText = varContent;											
									}
									break;

								case "Points as of date text (POS formatted":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryPointsAsOfLabel = spreadsheet["B"+idx].w.replace(/(\w\w\W\w\w\W\w\w\w\w|\w\w\w\w\W\w\w\W\w\w|\w\W\w\W\w)/, "<b>${AsOfDate}</b>");
									} else {
										obj.rewardsObj.summaryPointsAsOfLabel = varContent.replace(/(\w\w\W\w\w\W\w\w\w\w|\w\w\w\w\W\w\w\W\w\w|\w\W\w\W\w)/, "<b>${AsOfDate}</b>");											
									}
									break;

								case "Available text":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryAvailablePtsLabel = spreadsheet["B"+idx].w.split(" ")[0];
									} else {
										obj.rewardsObj.summaryAvailablePtsLabel = varContent.split(" ")[0];
									}
									break;

								case "Pending text":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryPendingPtsLabel = spreadsheet["B"+idx].w.split(" ")[0];
									} else {
										obj.rewardsObj.summaryPendingPtsLabel = varContent.split(" ")[0];											
									}
									break;

								case ">= pts text":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryEnoughPtsLabel = spreadsheet["B"+idx].w;
									} else {
										obj.rewardsObj.summaryEnoughPtsLabel = varContent;											
									}
									break;

								case "Hotel nights label (top bar)":
									obj.rewardsObj.hotelNights = {};
									obj.rewardsObj.dollarSpent = {};
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryNightsSpentLabel = spreadsheet["B"+idx].w;
									} else {
										obj.rewardsObj.summaryNightsSpentLabel = varContent;											
									}
									break;

								case "CTA (Redeem now) text":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryRedeemLinkText = spreadsheet["B"+idx].w + "&nbspc;&raquo;";
									} else {
										obj.rewardsObj.summaryRedeemLinkText = varContent + "&nbspc;&raquo;";											
									}
									break;

								case "Dollars spent label (bottom bar)":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryDollarsSpentLabel = spreadsheet["B"+idx].w;
									} else {
										obj.rewardsObj.summaryDollarsSpentLabel = varContent;											
									}
									break;

								case "< 3500 pts text":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryNeededPtsLabel = spreadsheet["B"+idx].w;
									} else {
										obj.rewardsObj.summaryNeededPtsLabel = varContent;											
									}
									break;

								case "See activity text":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryActivityLinkText = spreadsheet["B"+idx].w + "&nbspc;&raquo;";
									} else {
										obj.rewardsObj.summaryActivityLinkText = varContent + "&nbspc;&raquo;";											
									}
									break;

								case "CTA (Book now) text":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.summaryBookLinkText = spreadsheet["B"+idx].w + "&nbspc;&raquo;";
									} else {
										obj.rewardsObj.summaryBookLinkText = varContent + "&nbspc;&raquo;";											
									}
									break;

								case "Top bar value: silver":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.hotelNights.silverPoints = parseInt(spreadsheet["B"+idx].w.match(/[\d]/)[0]);
										obj.rewardsObj.hotelNights.goldPoints = parseInt(spreadsheet["B"+(idx+1)].w.match(/[\d][\d]/)[0]);
										obj.rewardsObj.hotelNights.silverThresholdText = spreadsheet["B"+idx].w.match(/[\d]/)[0];
										obj.rewardsObj.hotelNights.goldThresholdText = spreadsheet["B"+(idx+1)].w.match(/[\d][\d]/)[0];
										obj.rewardsObj.hotelNights.silverPointsText = spreadsheet["B"+idx].w.match(/[\D]+/)[0].trim();
										obj.rewardsObj.hotelNights.goldPointsText = spreadsheet["B"+(idx+1)].w.match(/[^\d][^\d]*/)[0].trim();
									} else {
										obj.rewardsObj.hotelNights.silverPoints = parseInt(varContent.match(/[\d]/)[0]);
										obj.rewardsObj.hotelNights.goldPoints = parseInt(spreadsheet[column+(idx+1)].w.match(/[\d][\d]/)[0]);											
										obj.rewardsObj.hotelNights.silverThresholdText = varContent.match(/[\d]/)[0];
										obj.rewardsObj.hotelNights.goldThresholdText = spreadsheet[column+(idx+1)].w.match(/[\d][\d]/)[0];
										obj.rewardsObj.hotelNights.silverPointsText = spreadsheet[column+idx].w.match(/[\D]+/)[0].trim();
										obj.rewardsObj.hotelNights.goldPointsText = spreadsheet[column+(idx+1)].w.match(/[^\d][^\d]*/)[0].trim();											
									}
									break;

								case "Bottom bar value: silver":
									obj.rewardsObj.dollarSpent.silverThresholdText = varContent.match(/(\d\d\d.\d\d\d.\d\d\d|\d\d.\d\d\d.\d\d\d|\d.\d\d\d.\d\d\d|\d\d\d.\d\d\d|\d\d.\d\d\d|\d.\d\d\d|\d\d\d|\d\d|\d)/)[0];
									obj.rewardsObj.dollarSpent.goldThresholdText = spreadsheet[column+(idx+1).toString()].w.match(/(\d\d\d.\d\d\d.\d\d\d|\d\d.\d\d\d.\d\d\d|\d.\d\d\d.\d\d\d|\d\d\d.\d\d\d|\d\d.\d\d\d|\d.\d\d\d|\d\d\d|\d\d|\d)/)[0];
									break;

								case "Progress bar tier label: +blue":
									if (varContent === "same as US EN 1033") {
										obj.rewardsObj.hotelNights.blueLabel = spreadsheet["B"+idx].w;
										obj.rewardsObj.hotelNights.silverLabel = spreadsheet["B"+(idx+1)].w;
										obj.rewardsObj.hotelNights.goldLabel = spreadsheet["B"+(idx+2)].w;
										obj.rewardsObj.dollarSpent.blueLabel = spreadsheet["B"+idx].w;
										obj.rewardsObj.dollarSpent.silverLabel = spreadsheet["B"+(idx+1)].w;
										obj.rewardsObj.dollarSpent.goldLabel = spreadsheet["B"+(idx+2)].w;
									} else {
										obj.rewardsObj.hotelNights.blueLabel = varContent;
										obj.rewardsObj.hotelNights.silverLabel = spreadsheet[column+(idx+1)].w;
										obj.rewardsObj.hotelNights.goldLabel = spreadsheet[column+(idx+2)].w;
										obj.rewardsObj.dollarSpent.blueLabel = varContent;
										obj.rewardsObj.dollarSpent.silverLabel = spreadsheet[column+(idx+1)].w;
										obj.rewardsObj.dollarSpent.goldLabel = spreadsheet[column+(idx+2)].w;
										obj.rewardsObj.dollarSpent.preCurrencySymbol = spreadsheet[column+"6"].w;											
									}
									break;

								default:
									//Do nothing
							}	
						}
						//If the cell does not have data then we don't need to add it to the loc file, so we do nothing
					}
					console.log("	- RWM Module completed...");
					break;

				//Construction of the punchcard module. This can be easily constructed withoud a massive switch statement.
				case "Punchcard Progress Module":
					obj.locVars.punchcard_progress_module = {};
					//This is formatting and variable insertion for the NightsStayed variable, since it is the same for all three types of recipients (eligible, registered, and completed)
					var nightsStayed = spreadsheet[column+(idx+6)].w.replace(/[X]/, "${PunchCardNights}");
					nightsStayed = nightsStayed.replace(/\u005B..........\u005D/, "${AsOfDate}");
					obj.locVars.punchcard_progress_module.NightsStayed = nightsStayed;
					//Now we begin looping through the rows, looking for the targeting describing rows
					for (var idx = lowerBound; idx <= upperBound; idx++) {
						var varName = spreadsheet["A"+idx].w;
						//When we find one, we store variables for that recipient type					
						if (varName.match("MER")) {
							if (varName.match("CER")) {
								if (varName.match("COMPLETED")) {
									//If they are registered and completed
									obj.locVars.punchcard_progress_module.Punchcard_header_full = spreadsheet[column+(idx+1)].w;
									obj.locVars.punchcard_progress_module.Punchcard_body_full = spreadsheet[column+(idx+2)].w;
									obj.locVars.punchcard_progress_module.Punchcard_bodyText_full = spreadsheet[column+(idx+3)].w;
									obj.locVars.punchcard_progress_module.Redeem_now_link = spreadsheet[column+(idx+5)].w + "&";
								} else {
									//If they are registered but not completed
									obj.locVars.punchcard_progress_module.Punchcard_header_registered = spreadsheet[column+(idx+1)].w;
									obj.locVars.punchcard_progress_module.Punchcard_body_registered = spreadsheet[column+(idx+2)].w;
									obj.locVars.punchcard_progress_module.Punchcard_bodyText_registered = spreadsheet[column+(idx+3)].w;
									obj.locVars.punchcard_progress_module.See_hotels_link = spreadsheet[column+(idx+5)].w + "&";									
								}
							} else {
								//If they are eligible but have not registered yet
								obj.locVars.punchcard_progress_module.Punchcard_header_register = spreadsheet[column+(idx+1)].w;
								obj.locVars.punchcard_progress_module.Punchcard_body_register = spreadsheet[column+(idx+2)].w;
								obj.locVars.punchcard_progress_module.Punchcard_bodyText_register = spreadsheet[column+(idx+3)].w;
								obj.locVars.punchcard_progress_module.Register_now_link = spreadsheet[column+(idx+5)].w + "&";	
							}
						}
					}
					console.log("	- Punchcard Module completed...");
					break;

				//Construction of the ELE benefits module. This module requires a large amount of string manipulation.
				case "ELE Benefits Reminder Module":
					obj.locVars.ele_benefits_reminder_module = {};
					for (var idx = lowerBound; idx <= upperBound; idx++) {
						var currentVar = spreadsheet["A"+idx.toString()].w;
						//Most of this wall of code is just text manipulation. Replacing, adding, and removing things to make this compatible with the current configuration of the modules and campaign body.
						if (currentVar === "Header") {
							if (spreadsheet[column+idx].w === "same as US EN 1033") {
								obj.locVars.ele_benefits_reminder_module.ELE_header = spreadsheet["B"+idx].w;
							}
							obj.locVars.ele_benefits_reminder_module.ELE_header = spreadsheet[column+idx].w;
						}
						//If we are at place 1
						if (currentVar.match("V1")) {
							//Check for this case again
							if (spreadsheet[column+idx].w === "same as AU EN 3081") {
								obj.locVars.ele_benefits_reminder_module.ELE_subheader_place1 = spreadsheet["U"+idx].w;
								var fullText = spreadsheet["U"+(idx+1)].w;
								var strongText = spreadsheet["U"+(idx+2)].w;									
							} else if (spreadsheet[column+idx].w === "same as US EN 1033") {
								obj.locVars.ele_benefits_reminder_module.ELE_subheader_place1 = spreadsheet["B"+idx].w;
								var fullText = spreadsheet["B"+(idx+1)].w;
								var strongText = spreadsheet["B"+(idx+2)].w;	
							} else {
								obj.locVars.ele_benefits_reminder_module.ELE_subheader_place1 = spreadsheet[column+idx].w;
								var fullText = spreadsheet[column+(idx+1)].w;
								var strongText = spreadsheet[column+(idx+2)].w;	
							}
							//These strings go through a few variables as the string gets split into smaller and smaller pieces. It saves what is needed as soon as it is isolated.
							if (strongText == "N/A") {
								var splitStrongText = fullText.split(/\.|。/);
							} else {
								var splitStrongText = fullText.split(strongText);								
							}
							//This first variable, the strongPreText is the same for both places, so we only need to save it for this first one.
							obj.locVars.ele_benefits_reminder_module.ELE_strongPreText = splitStrongText[0];
							obj.locVars.ele_benefits_reminder_module.ELE_strongPostText = splitStrongText[1];							
							obj.locVars.ele_benefits_reminder_module.ELE_strongText__place1 = strongText;
							if (splitStrongText.length > 1) {
								var splitPhoneText = splitStrongText[1].split(/\.|。/);		
								//freaking thai....no punctuation and it doesn't want to split properly.						
							} else {
								var splitPhoneText = splitStrongText[0].split(/\.|。/);								
							}
							//Pull out the phone number and store it
							var phoneNumber = fullText.match(/(\d.\d\d\d.\d\d\d.\d\d\d\d|\d..\d\d\d..\d\d\d.\d\d\d\d|.\d.\d\d\d.\d\d\d.\d\d\d\d.)/)[0];
							//Split the phone sentence via the phone number that we just pulled out
							//FOR th_TH, YOU MUST MANUALLY GO IN TO THE LOC AND ASSIGN THE PREPHONE AND POST PHONE VARIABLES!!! THERE IS NO PUNCTUATION TO SPLIT BY.
							var splitPhoneComponents = ((splitPhoneText[1]) ? splitPhoneText[1].split(phoneNumber) : "MUST FIX MANUALLY");
							//Another variable that only needs to be saved once
							obj.locVars.ele_benefits_reminder_module.ELE_bodyPrePhone = splitPhoneComponents[0];
							//Now we save the phone number link (adding a little thingy to the beginning)
							obj.locVars.ele_benefits_reminder_module.ELE_bodyPhoneLink__place1 = "tel:+"+phoneNumber;
							obj.locVars.ele_benefits_reminder_module.ELE_bodyPhone__place1 = phoneNumber;
							obj.locVars.ele_benefits_reminder_module.ELE_bodyPostPhone__place1 = splitPhoneComponents[1]+".";
							//If we are at place 2
						} else if (currentVar.match("V2")) {
							//Check for this stupid case again
							if (spreadsheet[column+idx].w === "same as AU EN 3081") {
								obj.locVars.ele_benefits_reminder_module.ELE_subheader_place2 = spreadsheet["U"+idx].w;
								var fullText = spreadsheet["U"+(idx+1)].w;
								var strongText = spreadsheet["U"+(idx+2)].w;									
							} else if (spreadsheet[column+idx].w === "same as US EN 1033") {
								obj.locVars.ele_benefits_reminder_module.ELE_subheader_place2 = spreadsheet["B"+idx].w;
								var fullText = spreadsheet["B"+(idx+1)].w;
								var strongText = spreadsheet["B"+(idx+2)].w;	
							} else {
								obj.locVars.ele_benefits_reminder_module.ELE_subheader_place2 = spreadsheet[column+idx].w;
								var fullText = spreadsheet[column+(idx+1)].w;
								var strongText = spreadsheet[column+(idx+2)].w;	
							}
							//These strings go through a few variables as the string gets split into smaller and smaller pieces. It saves what is needed as soon as it is isolated.
							if (strongText == "N/A") {
								var splitStrongText = fullText.split(/\.|。/);
							} else {
								var splitStrongText = fullText.split(strongText);								
							}
							//This first variable, the strongPreText is the same for both places, so we only need to save it for this first one.					
							obj.locVars.ele_benefits_reminder_module.ELE_strongText__place2 = strongText;
							if (splitStrongText.length > 1) {
								var splitPhoneText = splitStrongText[1].split(/\.|。/);		
								//freaking thai....no punctuation and it doesn't want to split properly.						
							} else {
								var splitPhoneText = splitStrongText[0].split(/\.|。/);								
							}
							obj.locVars.ele_benefits_reminder_module.ELE_strongPostText = splitStrongText[1];
							//Pull out the phone number and store it.
							var phoneNumber = fullText.match(/(\d.\d\d\d.\d\d\d.\d\d\d\d|\d..\d\d\d..\d\d\d.\d\d\d\d|.\d.\d\d\d.\d\d\d.\d\d\d\d.)/)[0];				
							//Split the phone sentence via the phone number that we just pulled out
							var splitPhoneComponents = ((splitPhoneText[1]) ? splitPhoneText[1].split(phoneNumber) : splitPhoneText[0].split(phoneNumber));
							//Now we save the phone number link (adding a little thingy to the beginning)
							obj.locVars.ele_benefits_reminder_module.ELE_bodyPhoneLink__place2 = "tel:"+phoneNumber;
							obj.locVars.ele_benefits_reminder_module.ELE_bodyPhone__place2 = phoneNumber;
							obj.locVars.ele_benefits_reminder_module.ELE_bodyPostPhone__place2 = splitPhoneComponents[1]+".";
							//Another variable that only needs to be saved once, but this one is down here because place 2 is physically closer to this cell on the spreadsheet
							obj.locVars.ele_benefits_reminder_module.ELE_bodyPostPhone = splitPhoneComponents[1] + ".";
							obj.locVars.ele_benefits_reminder_module.ELE_bodyLinkText = spreadsheet[column+(idx+3)].w;
						}
					}	
					console.log("	- ELE Module completed...");
					break;

				//Construction of the citi card module.
				//AT THIS TIME, CITI IS APPLICABLE FOR US ONLY AND IS HARDCODED WITH NO LOCALE FILE PRESENCE
				case "Citi Cobrand acquisition banner":
					//code
					console.log("	- Citi module skipped...");
					break;

				//Construction of the spotlight module.
				case "Spotlight module 1 - PWP T&C notice":
					obj.locVars.spotlight1_module = {};			
					for (var idx = lowerBound; idx <= upperBound; idx++) {
						//Make sure there is data in the cell before entering the switch statement
						if (spreadsheet[column+idx]) {
							var varContent = spreadsheet[column+idx].w;
							switch (spreadsheet["A"+idx].v) {
								//Each case here represents a different row and a different variable. We use the row label instead of the row number so that this will still work if the order of the rows is changed.
								case "Icon":
									obj.locVars.spotlight1_module.spotlight_icon = varContent;
									break;

								case "Header":
									obj.locVars.spotlight1_module.spotlight_header = varContent;
									break;

								case "Body Copy":
									obj.locVars.spotlight1_module.spotlight_body = varContent;
									break;

								case "CTA":
									obj.locVars.spotlight1_module.spotlight_linkText = varContent;
									break;

								case "URL":
									obj.locVars.spotlight1_module.spotlight_link = varContent.trim();
									break;

								default:
									console.log("Data not handled, omitting...");
							}
						}
					}
					// console.log(obj.locVars.spotlight1_module);
					console.log("	- Spotlight module completed");
					break;					

				//Construction of the special offers module
				case "Special Offers Module":
					obj.locVars.special_offers_module = {};
					//Immediately store the header
					obj.locVars.special_offers_module.SpecialOffers_header = ((spreadsheet[column+(lowerBound)].w === "same as MX ES 2058") ? spreadsheet["E"+(lowerBound)].w : spreadsheet[column+(lowerBound)].w);
					var offers = [];
					//Here we are detecting which offers are present, initializing empty variables for them, and storing the order of the offers in an array
					for (var x = 1; x <= 5; x++) {
						if (spreadsheet["A"+(lowerBound+x)].w.match(/Special Offer [0-9]/i)) {
							//Store the name of the offer in an array for later use
							offers.push(spreadsheet[column+(lowerBound+x)].w.toLowerCase());
							//Prepare the variable names
							var iconLabel = "SpecialOffers_icon__offer" + x;
							var textLabel = "SpecialOffers_text__offer" + x;
							var linkTextLabel = "SpecialOffers_linkText__offer" + x;
							var urlLabel = "SpecialOffers_URL__offer" + x;
							//Create variables for each offer (each offer has the same format, just different content)
							obj.locVars.special_offers_module[iconLabel] = "N/A";
							obj.locVars.special_offers_module[textLabel] = "N/A";
							obj.locVars.special_offers_module[linkTextLabel] = "N/A";
							obj.locVars.special_offers_module[urlLabel] = "N/A";
						}
					}
					// console.log(offers.length.toString(), "special offers found.");
					// console.log(offers);					
					for (var idx = lowerBound; idx <= upperBound; idx++) {
						//Make sure there is data in the cell before entering the switch statement
						if (spreadsheet[column+idx]) {
							//makes the code a bit less messy. This is purely a replacement variable so I don't have to keep writing spreadsheet[column+idx].w
							var varName = spreadsheet["A"+idx].w.toLowerCase();

							if (offers.indexOf(varName) >= 0) {
								// console.log("Matched", varName, "at row", idx);
							}
							//Begin the switch statement that will store the variables for each module present. If we find that a module is not present then we will break.
							switch (varName) {

								case offers[0]:
									// console.log("Found offer:", offers[0]);
									//Same logic as before when handling this cell content, but this is inline for neatness.
									obj.locVars.special_offers_module.SpecialOffers_icon__offer1 = ((spreadsheet[column+(idx+1)].w === "same as MX ES 2058") ? spreadsheet["E"+(idx+1)].w : spreadsheet[column+(idx+1)].w);
									//Need to parse through the body text to pull out the link.
									var fullBodyContent = ((spreadsheet[column+(idx+2)] === "same as MX ES 2058") ? spreadsheet["E"+(idx+2)].w : spreadsheet[column+(idx+2)].w);
									//Link is always preceeded by a ! or .
									fullBodyContent = fullBodyContent.split(/(!|\.)/);
									//Link is always the last item in the array
									obj.locVars.special_offers_module.SpecialOffers_linkText__offer1 = fullBodyContent[fullBodyContent.length-1];
									//Remove the last item since we have already stored it, and then we will join the rest of the array to reconstruct the body text
									fullBodyContent.length -= 1;
									obj.locVars.special_offers_module.SpecialOffers_text__offer1 = fullBodyContent.join("");			
									obj.locVars.special_offers_module.SpecialOffers_URL__offer1 = ((spreadsheet[column+(idx+3)].w === "same as MX ES 2058") ? spreadsheet["E"+idx+3].w : spreadsheet[column+(idx+3)].w);
									obj.locVars.special_offers_module.SpecialOffers_URL__offer1 += "&";
									break;

								case offers[1]:
									// console.log("Found offer:", offers[1]);
									//Same logic as before when handling this cell content, but this is inline for neatness.
									obj.locVars.special_offers_module.SpecialOffers_icon__offer2 = ((spreadsheet[column+(idx+1)].w === "same as MX ES 2058") ? spreadsheet["E"+(idx+1)].w : spreadsheet[column+(idx+1)].w);
									//Need to parse through the body text to pull out the link.
									if (spreadsheet[column+(idx+2)] === "same as MX ES 2058") {
										var fullBodyContent = spreadsheet["E"+(idx+2)].w;
									} else if (spreadsheet[column+(idx+2)] === "same as US EN 1033") {
										var fullBodyContent = spreadsheet["B"+(idx+2)].w;
									} else if (spreadsheet[column+(idx+2)] === "same as AU EN 3081") {
										var fullBodyContent = spreadsheet["U"+(idx+2)].w;
									} else {
										var fullBodyContent = spreadsheet[column+(idx+2)].w;
									}
									//Link is always preceeded by a ! or .
									fullBodyContent = fullBodyContent.split(/(!|\.)/);
									//Link is always the last item in the array
									obj.locVars.special_offers_module.SpecialOffers_linkText__offer2 = fullBodyContent[fullBodyContent.length-1];
									//Remove the last item since we have already stored it, and then we will join the rest of the array to reconstruct the body text
									fullBodyContent.length -= 1;
									obj.locVars.special_offers_module.SpecialOffers_text__offer2 = fullBodyContent.join("");			
									obj.locVars.special_offers_module.SpecialOffers_URL__offer2 = ((spreadsheet[column+(idx+3)].w === "same as MX ES 2058") ? spreadsheet["E"+idx+3].w : spreadsheet[column+(idx+3)].w);
									obj.locVars.special_offers_module.SpecialOffers_URL__offer2 += "&";
									break;

								case offers[2]:
									// console.log("Found offer:", offers[2]);
									//Same logic as before when handling this cell content, but this is inline for neatness.
									obj.locVars.special_offers_module.SpecialOffers_icon__offer3 = ((spreadsheet[column+(idx+1)].w === "same as MX ES 2058") ? spreadsheet["E"+(idx+1)].w : spreadsheet[column+(idx+1)].w);
									//Need to parse through the body text to pull out the link.
									var fullBodyContent = ((spreadsheet[column+(idx+2)] === "same as MX ES 2058") ? spreadsheet["E"+(idx+2)].w : spreadsheet[column+(idx+2)].w);
									//Link is always preceeded by a ! or .
									fullBodyContent = fullBodyContent.split(/(!|\.)/);
									//Link is always the last item in the array
									obj.locVars.special_offers_module.SpecialOffers_linkText__offer3 = fullBodyContent[fullBodyContent.length-1];
									//Remove the last item since we have already stored it, and then we will join the rest of the array to reconstruct the body text
									fullBodyContent.length -= 1;
									obj.locVars.special_offers_module.SpecialOffers_text__offer3 = fullBodyContent.join("");			
									obj.locVars.special_offers_module.SpecialOffers_URL__offer3 = ((spreadsheet[column+(idx+3)].w === "same as MX ES 2058") ? spreadsheet["E"+idx+3].w : spreadsheet[column+(idx+3)].w);
									obj.locVars.special_offers_module.SpecialOffers_URL__offer3+= "&";
									break;

								case offers[3]:
									// console.log("Found offer:", offers[3]);									
									//Same logic as before when handling this cell content, but this is inline for neatness.
									obj.locVars.special_offers_module.SpecialOffers_icon__offer4 = ((spreadsheet[column+(idx+1)].w === "same as MX ES 2058") ? spreadsheet["E"+(idx+1)].w : spreadsheet[column+(idx+1)].w);
									//Need to parse through the body text to pull out the link.
									var fullBodyContent = ((spreadsheet[column+(idx+2)] === "same as MX ES 2058") ? spreadsheet["E"+(idx+2)].w : spreadsheet[column+(idx+2)].w);
									//Link is always preceeded by a ! or .
									fullBodyContent = fullBodyContent.split(/(!|\.)/);
									//Link is always the last item in the array
									obj.locVars.special_offers_module.SpecialOffers_linkText__offer4 = fullBodyContent[fullBodyContent.length-1];
									//Remove the last item since we have already stored it, and then we will join the rest of the array to reconstruct the body text
									fullBodyContent.length -= 1;
									obj.locVars.special_offers_module.SpecialOffers_text__offer4 = fullBodyContent.join("");			
									obj.locVars.special_offers_module.SpecialOffers_URL__offer4 = ((spreadsheet[column+(idx+3)].w === "same as MX ES 2058") ? spreadsheet["E"+idx+3].w : spreadsheet[column+(idx+3)].w);
									obj.locVars.special_offers_module.SpecialOffers_URL__offer4 += "&";
									break;

								case offers[4]:
									// console.log("Found offer:", offers[4]);
									//Same logic as before when handling this cell content, but this is inline for neatness.
									obj.locVars.special_offers_module.SpecialOffers_icon__offer5 = ((spreadsheet[column+(idx+1)].w === "same as MX ES 2058") ? spreadsheet["E"+(idx+1)].w : spreadsheet[column+(idx+1)].w);
									//Need to parse through the body text to pull out the link.
									var fullBodyContent = ((spreadsheet[column+(idx+2)] === "same as MX ES 2058") ? spreadsheet["E"+(idx+2)].w : spreadsheet[column+(idx+2)].w);
									//Link is always preceeded by a ! or .
									fullBodyContent = fullBodyContent.split(/(!|\.)/);
									//Link is always the last item in the array
									obj.locVars.special_offers_module.SpecialOffers_linkText__offer5 = fullBodyContent[fullBodyContent.length-1];
									//Remove the last item since we have already stored it, and then we will join the rest of the array to reconstruct the body text
									fullBodyContent.length -= 1;
									obj.locVars.special_offers_module.SpecialOffers_text__offer5 = fullBodyContent.join("");			
									obj.locVars.special_offers_module.SpecialOffers_URL__offer5 = ((spreadsheet[column+(idx+3)].w === "same as MX ES 2058") ? spreadsheet["E"+idx+3].w : spreadsheet[column+(idx+3)].w);
									obj.locVars.special_offers_module.SpecialOffers_URL__offer5 += "&";
									break;

								default:
									break;

							}
						}
					}		
					// console.log(obj.locVars.special_offers_module);		
					console.log("	- Special offers module completed...");		
					break;

				//This module is also en_US only and is hardcoded
				case "Expedia+ business acquisition banner":
					obj.locVars[modName] = {};						
					//code
					console.log("	- Business acq. modules skipped...");
					break;

				//Same as above module
				case "Expedia+ business acquisition banner - CER VERSION":
					obj.locVars[modName] = {};						
					//code
					break;

				//Construction of the education module. This one's a doosie...
				case "Education Module":
					obj.locVars.education_module = {};
					//Immediately store the headers, as they are right below the module header
					obj.locVars.education_module.benefits_header__blue = ((spreadsheet[column+(lowerBound)].w === "same as US EN 1033") ? spreadsheet["B"+(lowerBound)].w : spreadsheet[column+(lowerBound)].w);
					obj.locVars.education_module.benefits_header__silver = ((spreadsheet[column+(lowerBound+1)].w === "same as US EN 1033") ? spreadsheet["B"+(lowerBound+1)].w : spreadsheet[column+(lowerBound+1)].w);
					obj.locVars.education_module.benefits_header__gold = ((spreadsheet[column+(lowerBound+2)].w === "same as US EN 1033") ? spreadsheet["B"+(lowerBound+2)].w : spreadsheet[column+(lowerBound+1)].w);
					//This module is structured with three modules names left, right, and center. They each have the same structure and format.
					for (var idx = lowerBound; idx <= upperBound; idx++) {
						var varName = spreadsheet["A"+idx].w;
						if (varName === "Left Module") {
							//This massive wall of variable storage includes the "same as ...." logic and splitting to isolate strings. 
							obj.locVars.education_module.benefits_left_icon__blue = ((spreadsheet[column+(idx+1)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+1)].w : spreadsheet[column+(idx+1)].w);
							obj.locVars.education_module.benefits_left_text__blue = ((spreadsheet[column+(idx+2)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+2)].w.split(".")[0] : spreadsheet[column+(idx+2)].w.split(".")[0]);
							obj.locVars.education_module.benefits_left_linkText__blue = ((spreadsheet[column+(idx+3)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+3)].w : spreadsheet[column+(idx+3)].w);
							obj.locVars.education_module.benefits_left_linkURL__blue = "${edu_sign_in_to_shop_linkURL}";
							obj.locVars.education_module.benefits_left_linkHashText_blue = ((spreadsheet[column+(idx+4)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+4)].w.split("#")[1] : spreadsheet[column+(idx+4)].w.split("#")[1]);
							obj.locVars.education_module.benefits_left_icon__silver = ((spreadsheet[column+(idx+5)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+5)].w : spreadsheet[column+(idx+5)].w);
							obj.locVars.education_module.benefits_left_text__silver = ((spreadsheet[column+(idx+6)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+6)].w.split(".")[0] : spreadsheet[column+(idx+6)].w.split(".")[0]);
							obj.locVars.education_module.benefits_left_linkText__silver = ((spreadsheet[column+(idx+7)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+7)].w : spreadsheet[column+(idx+7)].w);
							obj.locVars.education_module.benefits_left_linkURL__silver = "${edu_amenities_linkURL}";
							obj.locVars.education_module.benefits_left_linkHashText_silver = ((spreadsheet[column+(idx+8)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+8)].w.split("#")[1] : spreadsheet[column+(idx+8)].w.split("#")[1]);
							obj.locVars.education_module.benefits_left_icon__gold = ((spreadsheet[column+(idx+9)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+9)].w : spreadsheet[column+(idx+9)].w);
							obj.locVars.education_module.benefits_left_text__gold = ((spreadsheet[column+(idx+10)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+10)].w.split(".")[0] : spreadsheet[column+(idx+10)].w.split(".")[0]);
							obj.locVars.education_module.benefits_left_linkText__gold = ((spreadsheet[column+(idx+11)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+11)].w : spreadsheet[column+(idx+11)].w);
							obj.locVars.education_module.benefits_left_linkURL__gold = "${edu_amenities_linkURL}";
							obj.locVars.education_module.benefits_left_linkHashText_gold = ((spreadsheet[column+(idx+12)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+12)].w.split("#")[1] : spreadsheet[column+(idx+12)].w.split("#")[1]);

						} else if (varName === "Middle Module") {
							obj.locVars.education_module.benefits_middle_icon__blue = ((spreadsheet[column+(idx+1)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+1)].w : spreadsheet[column+(idx+1)].w);
							obj.locVars.education_module.benefits_middle_text__blue = ((spreadsheet[column+(idx+2)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+2)].w.split(".")[0] : spreadsheet[column+(idx+2)].w.split(".")[0]);
							obj.locVars.education_module.benefits_middle_linkText__blue = ((spreadsheet[column+(idx+3)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+3)].w : spreadsheet[column+(idx+3)].w);
							obj.locVars.education_module.benefits_middle_linkURL__blue = "${edu_sign_in_to_shop_linkURL}";
							obj.locVars.education_module.benefits_middle_linkHashText_blue = ((spreadsheet[column+(idx+4)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+4)].w.split("#")[1] : spreadsheet[column+(idx+4)].w.split("#")[1]);
							obj.locVars.education_module.benefits_middle_icon__silver = ((spreadsheet[column+(idx+5)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+5)].w : spreadsheet[column+(idx+5)].w);
							obj.locVars.education_module.benefits_middle_text__silver = ((spreadsheet[column+(idx+6)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+6)].w.split(".")[0] : spreadsheet[column+(idx+6)].w.split(".")[0]);
							obj.locVars.education_module.benefits_middle_linkText__silver = ((spreadsheet[column+(idx+7)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+7)].w : spreadsheet[column+(idx+7)].w);
							obj.locVars.education_module.benefits_middle_linkURL__silver = "${edu_amenities_linkURL}";
							obj.locVars.education_module.benefits_middle_linkHashText_silver = ((spreadsheet[column+(idx+8)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+8)].w.split("#")[1] : spreadsheet[column+(idx+8)].w.split("#")[1]);
							obj.locVars.education_module.benefits_middle_icon__gold = ((spreadsheet[column+(idx+9)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+9)].w : spreadsheet[column+(idx+9)].w);
							obj.locVars.education_module.benefits_middle_text__gold = ((spreadsheet[column+(idx+10)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+10)].w.split(".")[0] : spreadsheet[column+(idx+10)].w.split(".")[0]);
							obj.locVars.education_module.benefits_middle_linkText__gold = ((spreadsheet[column+(idx+11)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+11)].w : spreadsheet[column+(idx+11)].w);
							obj.locVars.education_module.benefits_middle_linkURL__gold = "${edu_amenities_linkURL}";
							obj.locVars.education_module.benefits_middle_linkHashText_gold = ((spreadsheet[column+(idx+12)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+12)].w.split("#")[1] : spreadsheet[column+(idx+12)].w.split("#")[1]);
						} else if (varName === "Right Module") {
							obj.locVars.education_module.benefits_right_icon__blue = ((spreadsheet[column+(idx+1)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+1)].w : spreadsheet[column+(idx+1)].w);
							obj.locVars.education_module.benefits_right_text__blue = ((spreadsheet[column+(idx+2)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+2)].w.split(".")[0] : spreadsheet[column+(idx+2)].w.split(".")[0]);
							obj.locVars.education_module.benefits_right_linkText__blue = ((spreadsheet[column+(idx+3)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+3)].w : spreadsheet[column+(idx+3)].w);
							obj.locVars.education_module.benefits_right_linkURL__blue = "${edu_sign_in_to_shop_linkURL}";
							obj.locVars.education_module.benefits_right_linkHashText_blue = ((spreadsheet[column+(idx+4)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+4)].w.split("#")[1] : spreadsheet[column+(idx+4)].w.split("#")[1]);
							obj.locVars.education_module.benefits_right_icon__silver = ((spreadsheet[column+(idx+5)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+5)].w : spreadsheet[column+(idx+5)].w);
							obj.locVars.education_module.benefits_right_text__silver = ((spreadsheet[column+(idx+6)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+6)].w.split(".")[0] : spreadsheet[column+(idx+6)].w.split(".")[0]);
							obj.locVars.education_module.benefits_right_linkText__silver = ((spreadsheet[column+(idx+7)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+7)].w : spreadsheet[column+(idx+7)].w);
							obj.locVars.education_module.benefits_right_linkURL__silver = "${edu_amenities_linkURL}";
							obj.locVars.education_module.benefits_right_linkHashText_silver = ((spreadsheet[column+(idx+8)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+8)].w.split("#")[1] : spreadsheet[column+(idx+8)].w.split("#")[1]);
							obj.locVars.education_module.benefits_right_icon__gold = ((spreadsheet[column+(idx+9)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+9)].w : spreadsheet[column+(idx+9)].w);
							obj.locVars.education_module.benefits_right_text__gold = ((spreadsheet[column+(idx+10)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+10)].w.split(".")[0] : spreadsheet[column+(idx+10)].w.split(".")[0]);
							obj.locVars.education_module.benefits_right_linkText__gold = ((spreadsheet[column+(idx+11)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+11)].w : spreadsheet[column+(idx+11)].w);
							obj.locVars.education_module.benefits_right_linkURL__gold = "${edu_amenities_linkURL}";
							obj.locVars.education_module.benefits_right_linkHashText_gold = ((spreadsheet[column+(idx+12)].w === "same as US EN 1033") ? spreadsheet["B"+(idx+12)].w.split("#")[1] : spreadsheet[column+(idx+12)].w.split("#")[1]);
						}
					}
					// console.log(obj.locVars.education_module)
					console.log("	- Education module completed...");
					break;

				case "Opt-In Module":
					obj.locVars.optin_module = {};						
					obj.locVars.optin_module.EmailSignup_header = ((spreadsheet[column+lowerBound].w === "same as US EN 1033") ? spreadsheet["B"+lowerBound].w : spreadsheet[column+lowerBound].w);
					//
					//Because of the way that the body text is put into the spreadsheet, there is not any way to isolate the link text from the rest of the body. This must still be hand coded :-/
					//
					console.log("	- Opt-in module completed...");
					break;
			}
		}
	}
	// console.log(obj);
	return obj;
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

//This function does many things, all related to further preparing the strings for insertion into a .ftl document. Some tasks include removing unnecessary html spans, replacing them with and adding <b> tags, replacing icons with their html entities, and adding any additional freemarker syntax.
function normalizeStrings(locales) {
	//Begin looping through to get to each element
	//Change this back to locales.length when finished with development
	for (var i = 0; i < 1; i++) {
		//Deal with locVars
		//Need nested for loops to access everything
		for (module in locales[i].content.locVars) {
			for (variable in locales[i].content.locVars[module]) {
				var thisVar = ((locales[i].content.locVars[module][variable] != undefined) ? locales[i].content.locVars[module][variable] : "N/A");
				//Search and replace for ampersands and the reg/trade symbols
				thisVar = thisVar.replace(/\u005B|\u005D/g, "");
				thisVar = thisVar.replace(/(\r|\n)/g, "");
				if (thisVar.match(/(&|®|™|)/) && !variable.match(/url/i)) {
					var amp = thisVar.match(/&/);
					var trade = thisVar.match(/™/);
					var reg = thisVar.match(/®/);
					if (amp) {
						thisVar = thisVar.replace("&", "&amp;");
					} else if (reg) {
						thisVar = thisVar.replace("®", "&reg;");
					} else if (trade) {
						thisVar = thisVar.replace("™", "&trade;");
					} 
					locales[i].content.locVars[module][variable] = thisVar;
				}
				//Formatting for link text
				if (variable.match(/linkText/i) && thisVar != "N/A") {
					//Remove ridiculous line breaks that are apparently coming out of nowhere...also spaces in front of text
					var end = thisVar.search(">");
					if (thisVar[0] === " " && end) {
						thisVar = thisVar.slice(1, end+1);						
					} else {
						thisVar = thisVar.slice(0, end+1);
					}
					//If there is no caret already at the end, add one
					if (!end) {
						//We still may need to slice a spacae off the front.
						if (thisVar[0] == " ") {
							thisVar = thisVar.slice(1, thisVar.length-1);
						}
						//Add the caret and space
						thisVar += "&nbspc;&raquo;";
					} else {
						//If end does exist, then we can replace the caret with the raquo thing.
						thisVar = thisVar.replace(/\s>$|>$/g, "");
						thisVar += " &nbspc;&raquo;";
					}	
					//This will be at the end of every case
					locales[i].content.locVars[module][variable] = thisVar;
				}
				//Remove unnecessary HTML
				if (variable.match(/strong/i) && !variable.match(/pre|post/i)) {
				}
			}
		}
		//Deal with rewardsObj
		var fileName = "./output/loc_" + locales[i].name + ".ftl";
		fs.appendFile(fileName, JSON.stringify(locales[i], null, '\t'), function(err) {
			if (err) {
				console.log(err);
			}
		});
	}
}

