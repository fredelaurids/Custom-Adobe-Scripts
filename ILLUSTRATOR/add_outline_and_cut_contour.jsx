function main() {
  // Get active document, and selection
  var idoc = app.activeDocument;
  var sel = idoc.selection[0];

  if (sel == null) {
    alert("Nothing selected.");
    return;
  }

  // Disable alerts
  app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

  // In points
  var offset = 3;

  // Outline fill color
  var outlineColor = new CMYKColor();
  outlineColor.cyan = 0;
  outlineColor.magenta = 0;
  outlineColor.yellow = 0;
  outlineColor.black = 100;

  // Cut contour stroke color
  var cutContourColor = new CMYKColor();
  cutContourColor.cyan = 0;
  cutContourColor.magenta = 100;
  cutContourColor.yellow = 0;
  cutContourColor.black = 0;



  // Offset current selection
  offsetPath(idoc, sel, offset * 0.5);
  // Select again after offseting
  sel = idoc.selection[0];
  // Move to layer Outline
  moveToLayer(idoc, sel, "CUT_CONTOUR");
  // Set fill color
  setGroupStrokeColor(sel, cutContourColor);



  // Offset current selection
  offsetPath(idoc, sel, offset * 0.5);
  // Select again after offseting
  sel = idoc.selection[0];
  // Move to layer Outline
  moveToLayer(idoc, sel, "Outline");
  // Set fill color
  setGroupFillColor(sel, outlineColor);

  // Redraw
  app.redraw();

  app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
}

function offsetPath(idoc, sel, offset) {
  // Define the offset effect
  var xmlstringOffsetPath =
    '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst value I jntp 0 "/></LiveEffect>'; // 0 = round, 1=bevel, 2=miter
  xmlstringOffsetPath = xmlstringOffsetPath.replace("value", offset);

  idoc.selection = null;

  var dup = sel.duplicate(sel, ElementPlacement.PLACEAFTER);
  try {
    dup.filled = false;
  } catch (e) {}

  // Select new path before applying effects
  dup.selected = true;

  // Apply offset path Effect
  dup.applyEffect(xmlstringOffsetPath);
  app.redraw();

  // expand appearance
  app.executeMenuCommand("expandStyle");
  app.executeMenuCommand("Live Pathfinder Add");
  app.executeMenuCommand("expandStyle");
}

function moveToLayer(idoc, item, layerName) {
  // Move to layer
  var _layer = idoc.layers.getByName(layerName);
  item.move(_layer, ElementPlacement.PLACEATEND);
}

function setGroupFillColor(group, newColor) {
  // Get all the new elements in the group in the Outline layer
  for (var i = 0; i < group.pathItems.length; i++) {
    var element = group.pathItems[i];
    element.filled = true;
    element.stroked = false;
    element.fillColor = newColor;
  }
}

function setGroupStrokeColor(group, newColor) {
    // Get all the new elements in the group in the Outline layer
    for (var i = 0; i < group.pathItems.length; i++) {
      var element = group.pathItems[i];
      element.stroked = true;
      element.filled = false;
      element.strokeColor = newColor;
    }
  }

main();
