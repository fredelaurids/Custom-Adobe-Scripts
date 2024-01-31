app.doScript(
	main,
	ScriptLanguage.JAVASCRIPT,
	undefined,
	UndoModes.entireScript,
	"Create Modular Grid"
);

function main() {
	var idoc = app.activeDocument;

	// Needed variables to calculate modular grid
	// Baseline increment
	// Baseline grid start
	// Top margin
	// Number of paragraph lines on a page

	// Get top margin of the page that is currently visible
	var marginTop = idoc.layoutWindows[0].activePage.marginPreferences.top;

	// Get baseline increment of document
	var baselineIncrement = idoc.gridPreferences.baselineDivision;

	// Get baseline start
	var baselineStart = idoc.gridPreferences.baselineStart;

	// Calculate gutter between modules
	var gutter = baselineIncrement * 2 - (baselineStart - marginTop);

	// Prompt for number of lines in total
	var numberOfLines = prompt("Number of lines", 0);

	// Calculate bottom margin
	var pageHeight = Math.ceil(idoc.layoutWindows[0].activePage.bounds[2]); // Get height of active page
	var marginBottom =
		pageHeight - ((numberOfLines - 1) * baselineIncrement + baselineStart);

	// Set bottom margin
	idoc.layoutWindows[0].activePage.marginPreferences.bottom = marginBottom;

	// Set default color to gray, then prompt for colors
	var activeGuideColor = UIColors.gray;
	openDialog();
}

function drawModularGrid() {
	drawGuide(100, 0);
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

function openDialog() {
	var dialog = app.dialogs.add({
		name: "Create Modular Grid",
	});
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

	with (dialog) {
		with (dialogColumns.add()) {
			// with(dialogRows.add()) {
			//     with(dialogColumns.add()) {
			//         with(borderPanels.add()) {
			//             staticTexts.add({
			//                 staticLabel: "Rows:",
			//                 minWidth: 60
			//             });
			//             var numRowsBox = realEditboxes.add({
			//                 editValue: 1
			//             });
			//             staticTexts.add({
			//                 staticLabel: "Gutter:"
			//             });
			//             var rowGutterBox = realEditboxes.add({
			//                 editValue: 0.125
			//             });
			//         }
			//     }
			// }
			// with(dialogRows.add()) {
			//     with(dialogColumns.add()) {
			//         with(borderPanels.add()) {
			//             staticTexts.add({
			//                 staticLabel: "Columns:"
			//             });
			//             var numColumnsBox = realEditboxes.add({
			//                 editValue: 1
			//             });
			//             staticTexts.add({
			//                 staticLabel: "Gutter:"
			//             });
			//             var columnGutterBox = realEditboxes.add({
			//                 editValue: 0.125
			//             });
			//         }
			//     }
			// }
			with (dialogRows.add()) {
				with (borderPanels.add()) {
					var colorMenu = dropdowns.add({
						stringList: colorStrings,
						selectedIndex: 0,
						minWidth: 259,
					});
				}
			}
		}
	}
	var returned = dialog.show();

	if (returned) {
		// numRows = parseInt(numRowsBox.editContents);
		// numColumns = parseInt(numColumnsBox.editContents);
		// rowGutter = parseFloat(rowGutterBox.editContents);
		// columnGutter = parseFloat(columnGutterBox.editContents);
		activeGuideColor = UIColors[colorMenu.stringList[colorMenu.selectedIndex]];
		drawModularGrid();
	}
}
