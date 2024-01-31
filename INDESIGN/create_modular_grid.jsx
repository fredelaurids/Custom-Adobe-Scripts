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

	// BASELINE VARIABLES
	// Get x-height of the font of the selected paragraph style
	var xHeight = getXHeight(paragraphStyle);
	// Calculate baseline increment in points
	var baselineIncrement = new UnitValue(paragraphStyle.leading, "pt");
	// Calculate baseline start
	var baselineStart = marginTop + xHeight;

	// DOCUMENT DIMENSION VARIABLES
	// Get document height
	var pageHeight = Math.ceil(doc.layoutWindows[0].activePage.bounds[2]);

	// Get maximum number of lines possible from margintop to bottom of page
	var maxNumberOfLines = Math.floor(
		(pageHeight - marginTop) / baselineIncrement.as("mm")
	);
	alert("Maximum number of lines within margins: " + maxNumberOfLines);

	// Calculate number of lines per module
	var linesPerModule = Math.floor(
		(maxNumberOfLines - (numberOfModules - 1)) / numberOfModules
	);
	alert("Number of lines per module: " + linesPerModule);

	if (linesPerModule < 1) {
		alert(
			"Not enough lines to create a modular grid. Lines per module evaluated to: " +
				linesPerModule
		);
		return;
	}

	// Calculate total number of lines that fits within the number of modules, plus the gutter between them
	var numberOfLines = linesPerModule * numberOfModules + (numberOfModules - 1);
	alert("Total lines: " + numberOfLines);

	// Calculate gutter between modules
	var gutter = baselineIncrement.as("mm") * 2 - (baselineStart - marginTop);

	// Calculate bottom margin
	var marginBottom =
		pageHeight -
		((numberOfLines - 1) * baselineIncrement.as("mm") + baselineStart);

	// Set margins
	doc.layoutWindows[0].activePage.marginPreferences.top = marginTop;
	doc.layoutWindows[0].activePage.marginPreferences.bottom = marginBottom;

	// Set baseline increment
	doc.gridPreferences.baselineDivision = baselineIncrement.value + "pt"; // Add "pt" string to make it points, for some reason UnitValue conversion isn't working for me
	doc.gridPreferences.baselineStart = baselineStart;

	// Set Basic Text Frame first baseline offset to x-height
	doc.objectStyles.itemByName(
		"[Basic Text Frame]"
	).textFramePreferences.firstBaselineOffset = FirstBaseline.X_HEIGHT;

	// Draw the guides, starting at the margin top
    var currentX = 0;
	for (var i = 0; i < (numberOfModules * 2); i++) {
        drawGuide(marginTop + currentX, 0);

        // If index is 1, it's a gutter, otherwise it's a module
        if(i % 2 == 1) {
            currentX += gutter;
            //alert("adding gutter");
        } else {
            currentX += (linesPerModule * baselineIncrement.as("mm")) - (baselineIncrement.as("mm") - xHeight);
            //alert("adding module");
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

	with (dialog) {
		with (dialogColumns.add()) {
			with (dialogRows.add()) {
				with (dialogColumns.add()) {
					with (borderPanels.add()) {
						staticTexts.add({
							staticLabel: "Number of modules:",
						});
						var numberOfModulesBox = realEditboxes.add({
							editValue: 5,
						});
						staticTexts.add({
							staticLabel: "Margin top:",
						});
						var marginTopBox = realEditboxes.add({
							editValue:
								app.activeDocument.layoutWindows[0].activePage.marginPreferences
									.top,
						});
					}
				}
			}
			with (dialogRows.add()) {
				with (borderPanels.add()) {
					var colorMenu = dropdowns.add({
						stringList: colorStrings,
						selectedIndex: 0,
						minWidth: 259,
					});
				}
			}
			with (dialogRows.add()) {
				with (borderPanels.add()) {
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
		numberOfModules = parseInt(numberOfModulesBox.editContents);

		marginTop = parseFloat(marginTopBox.editContents);

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
