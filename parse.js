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
//This is also where all the functions get called
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
		locales[i].eyebrow_linktext = spreadsheet[currentCol+"23"].v;
		locales[i].eyebrow_link = spreadsheet[currentCol+"24"].v;
		instance.nextColumn();
	}
}

//This function parses through the spreadsheet and stores variables in each module, does one whole locale at a time then returns. Executed as called by buildLocales.
function buildModules(column) {
	//Loop through each module, we will go through each row within here creating the module for each locale
	var allMods = [];
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
										obj.rewardsObj.dollarSpent.silverThresholdText = varContent.match(/(\d\d\d.\d\d\d.\d\d\d|\d\d.\d\d\d.\d\d\d|\d.\d\d\d.\d\d\d|\d\d\d.\d\d\d|\d\d.\d\d\d|\d.\d\d\d)/)[0];
										obj.rewardsObj.dollarSpent.goldThresholdText = spreadsheet[column+(idx+1).toString()].w.match(/(\d\d\d.\d\d\d.\d\d\d|\d\d.\d\d\d.\d\d\d|\d.\d\d\d.\d\d\d|\d\d\d.\d\d\d|\d\d.\d\d\d|\d.\d\d\d)/)[0];
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
						break;

					//Construction of the punchcard module. This can be easily constructed withoud a massive switch statement.
					case "Punchcard Progress Module":
						obj.locVars.punchcard_progress_module = {};
						//This is formatting and variable insertion for the NightsStayed variable, since it is the same for all three types of recipients (eligible, registered, and completed)
						var nightsStayed = spreadsheet[column+(idx+6).toString()].w.replace(/[X]/, "${PunchCardNights}");
						nightsStayed.replace(/\[..........\]/, "${AsOfDate}");
						obj.locVars.punchcard_progress_module.NightsStayed = nightsStayed;
						//Now we begin looping through the rows, looking for the targeting describing rows
						for (var idx = lowerBound; idx <= upperBound; idx++) {
							var varName = spreadsheet["A"+idx.toString()].w;
							//When we find one, we store variables for that recipient type					
							if (varName.match("MER")) {
								if (varName.match("CER")) {
									if (varName.match("COMPLETED")) {
										//If they are registered and completed
										obj.locVars.punchcard_progress_module.Punchcard_header_full = spreadsheet[column+(idx+1).toString()].w;
										obj.locVars.punchcard_progress_module.Punchcard_body_full = spreadsheet[column+(idx+2).toString()].w;
										obj.locVars.punchcard_progress_module.Punchcard_bodyText_full = spreadsheet[column+(idx+3).toString()].w;
										obj.locVars.punchcard_progress_module.Redeem_now_link = spreadsheet[column+(idx+5).toString()].w + "&";
									} else {
										//If they are registered but not completed
										obj.locVars.punchcard_progress_module.Punchcard_header_registered = spreadsheet[column+(idx+1).toString()].w;
										obj.locVars.punchcard_progress_module.Punchcard_body_registered = spreadsheet[column+(idx+2).toString()].w;
										obj.locVars.punchcard_progress_module.Punchcard_bodyText_registered = spreadsheet[column+(idx+3).toString()].w;
										obj.locVars.punchcard_progress_module.See_hotels_link = spreadsheet[column+(idx+5).toString()].w + "&";									
									}
								} else {
									//If they are eligible but have not registered yet
									obj.locVars.punchcard_progress_module.Punchcard_header_register = spreadsheet[column+(idx+1).toString()].w;
									obj.locVars.punchcard_progress_module.Punchcard_body_register = spreadsheet[column+(idx+2).toString()].w;
									obj.locVars.punchcard_progress_module.Punchcard_bodyText_register = spreadsheet[column+(idx+3).toString()].w;
									obj.locVars.punchcard_progress_module.Register_now_link = spreadsheet[column+(idx+5).toString()].w + "&";	
								}
							}
						}
						break;

					case "ELE Benefits Reminder Module":
						obj.locVars.ele_benefits_reminder_module = {};
						for (var idx = lowerBound; idx <= upperBound; idx++) {
							var currentVar = spreadsheet["A"+idx.toString()].w;
							//Most of this wall of code is just text manipulation. Replacing, adding, and removing things to make this compatible with the current configuration of the modules and campaign body.
							if (currentVar === "Header") {
								if (spreadsheet[column+idx].w = "same as US EN 1033") {
									obj.locVars.ele_benefits_reminder_module.ELE_header = spreadsheet["B"+idx].w;
								}
								obj.locVars.ele_benefits_reminder_module.ELE_header = spreadsheet[column+idx].w;
							}
							//If we are at place 1
							if (currentVar.match("V1")) {
								//Check for this case again
								if (spreadsheet[column+idx].w === "same as AU EN 3081") {
									obj.locVars.ele_benefits_reminder_module.ELE_subheader_place1 = spreadsheet["U"+idx].w;
									var fullText = spreadsheet["U"+(idx+1)].h.replace('span style="font-weight: bold;"', "b");									
								} else {
									obj.locVars.ele_benefits_reminder_module.ELE_subheader_place1 = spreadsheet[column+idx].w;
									var fullText = spreadsheet[column+(idx+1)].h.replace('span style="font-weight: bold;"', "b");	
								}
								fullText = fullText.replace(/[\/]span/, "b");
								fullText = fullText.replace('<span style="">', "");
								fullText = fullText.replace(/<[\/]span>/, "");
								fullText = fullText.replace(/<br\/>/, "");
								//These strings go through a few variables as the string gets split into smaller and smaller pieces. It saves what is needed as soon as it is isolated.
								var splitStrongText = fullText.split("<b>");
								//This first variable, the strongPreText is the same for both places, so we only need to save it for this first one.
								obj.locVars.ele_benefits_reminder_module.ELE_strongPreText = splitStrongText[0];
								obj.locVars.ele_benefits_reminder_module.ELE_strongText__place1 = splitStrongText[1];
								var splitPhoneText = splitStrongText[2].split(".");
								obj.locVars.ele_benefits_reminder_module.ELE_strongPostText = splitStrongText[1]+". ";
								//Pull out the phone number and store it
								var phoneNumber = splitPhoneText[1].match(/1([0-9]|-|\(|\)|\s)+/)[0];
								//Split the phone sentence via the phone number that we just pulled out
								var splitPhoneComponents = splitPhoneText[1].split(phoneNumber);
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
									var fullText = spreadsheet["U"+(idx+1)].h.replace('span style="font-weight: bold;"', "b");									
								} else {
									obj.locVars.ele_benefits_reminder_module.ELE_subheader_place2 = spreadsheet[column+idx].w;
									var fullText = spreadsheet[column+(idx+1)].h.replace('span style="font-weight: bold;"', "b");	
								}
								fullText = fullText.replace(/[\/]span/, "b");
								fullText = fullText.replace('<span style="">', "");
								fullText = fullText.replace(/<[\/]span>/, "");
								fullText = fullText.replace(/<br\/>/, "");
								var splitStrongText = fullText.split("<b>");
								obj.locVars.ele_benefits_reminder_module.ELE_strongText__place2 = splitStrongText[1];
								var splitPhoneText = splitStrongText[2].split(".");
								obj.locVars.ele_benefits_reminder_module.ELE_strongPostText = splitStrongText[1]+". ";
								//Pull out the phone number and store it.
								var phoneNumber = splitPhoneText[1].match(/1([0-9]|-|\(|\)|\s)+/)[0];
								//Split the phone sentence via the phone number that we just pulled out
								var splitPhoneComponents = splitPhoneText[1].split(phoneNumber);
								//Now we save the phone number link (adding a little thingy to the beginning)
								obj.locVars.ele_benefits_reminder_module.ELE_bodyPhoneLink__place2 = "tel:"+phoneNumber;
								obj.locVars.ele_benefits_reminder_module.ELE_bodyPhone__place2 = phoneNumber;
								obj.locVars.ele_benefits_reminder_module.ELE_bodyPostPhone__place2 = splitPhoneComponents[1]+".";
								//Another variable that only needs to be saved once, but this one is down here because place 2 is physically closer to this cell on the spreadsheet
								obj.locVars.ele_benefits_reminder_module.ELE_bodyPostPhone = splitPhoneComponents[1] + ".";
								obj.locVars.ele_benefits_reminder_module.ELE_bodyLinkText = spreadsheet[column+(idx+3)].w;
							}
						}	
						console.log(obj.locVars.ele_benefits_reminder_module);				
						//code
						break;

					case "Citi Cobrand acquisition banner":
						obj.locVars[modName] = {};						
						//code
						break;

					case "Spotlight module 1 - PWP T&C notice":
						obj.locVars[modName] = {};						
						//code
						break;					

					case "Special Offers Module":
						obj.locVars[modName] = {};						
						//code
						break;

					case "Expedia+ business acquisition banner":
						obj.locVars[modName] = {};						
						//code
						break;

					case "Expedia+ business acquisition banner - CER VERSION":
						obj.locVars[modName] = {};						
						//code
						break;

					case "Education Module":
						obj.locVars[modName] = {};						
						//code
						break;

					case "Opt-In Module":
						obj.locVars[modName] = {};						
						//code
						break;
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

