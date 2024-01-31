// app.doScript(
// 	main,
// 	ScriptLanguage.JAVASCRIPT,
// 	undefined,
// 	UndoModes.entireScript,
// 	"Create Modular Grid"
// );

main();

function main() {
	if (app.activeDocument.layoutWindows.length > 0) {
		openDialog();
	} else {
		alert("No valid active document. Aborting.");
	}
}

function drawModularGrid() {
	var doc = app.activeDocument;

	/*-----------------------------------------------
    ---------------CALCULATE VARIABLES---------------
    -----------------------------------------------*/

	// Get x-height of the font of the selected paragraph style
	var xHeight = getXHeight(paragraphStyle);
	// Calculate baseline increment in points
	var baselineIncrement = new UnitValue(paragraphStyle.leading, "pt");
	// Calculate baseline start
	var baselineStart = marginTop + xHeight;

	// Get document height
	var pageHeight = Math.ceil(doc.layoutWindows[0].activePage.bounds[2]);

	// Get width within margins
	var pageWidth =
		Math.ceil(doc.layoutWindows[0].activePage.bounds[3]) -
		marginLeft -
		marginRight;

	// Get maximum number of lines possible from margintop to bottom of page
	var maxNumberOfLines = Math.floor(
		(pageHeight - marginTop) / baselineIncrement.as("mm")
	);
	alert("Maximum number of lines within margins: " + maxNumberOfLines);

	// Calculate number of lines per Row
	var linesPerRow = Math.floor(
		(maxNumberOfLines - (numberOfRows - 1)) / numberOfRows
	);
	alert("Number of lines per Row: " + linesPerRow);

	if (linesPerRow < 1) {
		alert(
			"Not enough lines to create a modular grid. Lines per Row evaluated to: " +
				linesPerRow
		);
		return;
	}

	// Calculate total number of lines that fits within the number of Rows, plus the gutter between them
	var numberOfLines = linesPerRow * numberOfRows + (numberOfRows - 1);
	alert("Total lines: " + numberOfLines);

	// Calculate gutter between Rows
	var gutter = baselineIncrement.as("mm") * 2 - (baselineStart - marginTop);

	// Calculate bottom margin
	var marginBottom =
		pageHeight -
		((numberOfLines - 1) * baselineIncrement.as("mm") + baselineStart);

	// Calculate Row height
	var rowHeight =
		linesPerRow * baselineIncrement.as("mm") -
		(baselineIncrement.as("mm") - xHeight);

	// Calculate Column width
	var colWidth = (pageWidth - gutter * (numberOfCols - 1)) / numberOfCols;

	/*-----------------------------------------------
    ---------------SET DOCUMENT STUFF----------------
    -----------------------------------------------*/

	// Set margins
	doc.layoutWindows[0].activePage.marginPreferences.top = marginTop;
	doc.layoutWindows[0].activePage.marginPreferences.left = marginLeft;
	doc.layoutWindows[0].activePage.marginPreferences.right = marginRight;
	doc.layoutWindows[0].activePage.marginPreferences.bottom = marginBottom;

	// Set baseline increment
	doc.gridPreferences.baselineDivision = baselineIncrement.value + "pt"; // Add "pt" string to make it points, for some reason UnitValue conversion isn't working for me
	doc.gridPreferences.baselineStart = baselineStart;

	// Set Basic Text Frame first baseline offset to x-height
	doc.objectStyles.itemByName(
		"[Basic Text Frame]"
	).textFramePreferences.firstBaselineOffset = FirstBaseline.X_HEIGHT;

	/*-----------------------------------------------
    ---------------DRAW HORIZONTAL GUIDES------------
    -----------------------------------------------*/

	// Draw the horizontal guides, starting at the top margin
	var currentX = 0;
	for (var i = 0; i < numberOfRows * 2; i++) {
		drawGuide(marginTop + currentX, 0);

		// If index is 0 it's a Row. If index is 1, it's a gutter.
		if (i % 2 == 0) {
			currentX += rowHeight;
		} else {
			currentX += gutter;
		}
	}

	/*-----------------------------------------------
    ---------------DRAW VERTICLES GUIDES-------------
    -----------------------------------------------*/

	// Draw the horizontal guides, starting at the left margin
	var currentY = 0;
	for (var i = 0; i < numberOfCols * 2; i++) {
		drawGuide(marginLeft + currentY, 1);

		// If index is 0 it's a Row. If index is 1, it's a gutter.
		if (i % 2 == 0) {
			currentY += colWidth;
		} else {
			currentY += gutter;
		}
	}
}

function openDialog() {
	// Create dialog with name
	var dialog = app.dialogs.add({
		name: "Create Modular Grid",
	});

	// Guide color choices
	var colorStrings = [
		"lightBlue",
		"red",
		"green",
		"blue",
		"yellow",
		"magenta",
		"cyan",
		"gray",
		"black",
		"orange",
		"darkGreen",
		"teal",
		"tan",
		"brown",
		"violet",
		"gold",
		"darkBlue",
		"pink",
		"lavender",
		"brickRed",
	];

	// Get all the paragraph styles
	var paragraphStyles = app.activeDocument.paragraphStyles;
	var paragraphStyleStrings = [];

	for (var i = 0; i < paragraphStyles.length; i++) {
		paragraphStyleStrings.push(paragraphStyles[i].name);
	}

	if (paragraphStyleStrings.length == 0) {
		alert("No paragraph styles found.");
	}

	var margins = [
		new UnitValue(
			app.activeDocument.layoutWindows[0].activePage.marginPreferences.top,
			"mm"
		),
		new UnitValue(
			app.activeDocument.layoutWindows[0].activePage.marginPreferences.left,
			"mm"
		),
		new UnitValue(
			app.activeDocument.layoutWindows[0].activePage.marginPreferences.right,
			"mm"
		),
	];

	with (dialog) {
		with (dialogColumns.add()) {
			with (dialogRows.add()) {
				with (dialogColumns.add()) {
					with (borderPanels.add()) {
						staticTexts.add({
							staticLabel: "Rows:",
						});
						var numberOfRowsBox = integerEditboxes.add({
							editValue: 5,
							smallNudge: 1,
							minimumValue: 2,
							maximumValue: 50,
						});
						staticTexts.add({
							staticLabel: "Columns (can be 0):",
						});
						var numberOfColsBox = integerEditboxes.add({
							editValue: 5,
							smallNudge: 1,
							minimumValue: 0,
							maximumValue: 50,
						});
					}
				}
			}
			with (dialogRows.add()) {
				with (borderPanels.add()) {
					staticTexts.add({
						staticLabel: "Margin top:",
					});
					var marginTopBox = measurementEditboxes.add({
						smallNudge: 0.5,
						largeNudge: 1,
						editUnits: MeasurementUnits.MILLIMETERS,
						editValue: margins[0].as("pt"),
					});
					staticTexts.add({
						staticLabel: "Margin left:",
					});
					var marginLeftBox = measurementEditboxes.add({
						smallNudge: 0.5,
						largeNudge: 1,
						editUnits: MeasurementUnits.MILLIMETERS,
						editValue: margins[1].as("pt"),
					});
					staticTexts.add({
						staticLabel: "Margin right:",
					});
					var marginRightBox = measurementEditboxes.add({
						smallNudge: 0.5,
						largeNudge: 1,
						editUnits: MeasurementUnits.MILLIMETERS,
						editValue: margins[2].as("pt"),
					});
					staticTexts.add({
						staticLabel:
							"Bottom margin is variable, and based the amount of total lines in the modular grid.",
					});
				}
			}
			with (dialogRows.add()) {
				with (borderPanels.add()) {
					staticTexts.add({
						staticLabel: "Guide color:",
					});
					var colorMenu = dropdowns.add({
						stringList: colorStrings,
						selectedIndex: 0,
						minWidth: 259,
					});
				}
			}
			with (dialogRows.add()) {
				with (borderPanels.add()) {
					staticTexts.add({
						staticLabel: "Base on paragraph style:",
					});
					var paragraphStyleMenu = dropdowns.add({
						stringList: paragraphStyleStrings,
						selectedIndex: 0,
						minWidth: 259,
					});
				}
			}
		}
	}
	var returned = dialog.show();

	if (returned) {
		numberOfRows = parseInt(numberOfRowsBox.editValue);
		numberOfCols = parseInt(numberOfColsBox.editValue);

		marginTop = parseFloat(new UnitValue(marginTopBox.editValue, "pt").as("mm"));
		marginLeft = parseFloat(new UnitValue(marginLeftBox.editValue, "pt").as("mm"));
		marginRight = parseFloat(new UnitValue(marginRightBox.editValue, "pt").as("mm"));

		paragraphStyle = paragraphStyles.itemByName(
			paragraphStyleMenu.stringList[paragraphStyleMenu.selectedIndex]
		);

		activeGuideColor = UIColors[colorMenu.stringList[colorMenu.selectedIndex]];
		drawModularGrid();
	}
}

function getXHeight(paragraphStyle) {
	var doc = app.activeDocument;

	var tempFrame = doc.textFrames.add({
		geometricBounds: doc.pages[0].bounds,
		contents: "x",
	});
	tempFrame.paragraphs.everyItem().appliedParagraphStyle = paragraphStyle;

	tempFrame.textFramePreferences.properties = {
		firstBaselineOffset: FirstBaseline.X_HEIGHT,
		minimumFirstBaselineOffset: 0,
		insetSpacing: 0,
		ignoreWrap: true,
		verticalJustification: VerticalJustification.TOP_ALIGN,
	};

	tempFrame.texts[0].alignToBaseline = false;

	var height =
		tempFrame.insertionPoints[0].baseline - tempFrame.geometricBounds[0];

	tempFrame.remove();

	return height;
}

function drawGuide(myGuideLocation, myGuideOrientation) {
	// Orientation 0 is horizontal, 1 is vertical
	with (app.activeWindow.activePage) {
		//Place guides at the margins of the page.
		if (myGuideOrientation == 0) {
			guides.add(undefined, {
				orientation: HorizontalOrVertical.horizontal,
				location: myGuideLocation,
				guideColor: activeGuideColor,
			});
		} else {
			guides.add(undefined, {
				orientation: HorizontalOrVertical.vertical,
				location: myGuideLocation,
				guideColor: activeGuideColor,
			});
		}
	}
}
