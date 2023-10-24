function main() {
    // Get active document, and selection
    var idoc = app.activeDocument;
    var sel = idoc.selection[0];

    if(sel == null) {
        alert("Nothing selected.");
        return;
    }

    // Disable alerts
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    // In points
    var offset = 3;

    // New fill
    var cmykColor = new CMYKColor();
    cmykColor.cyan = 0;
    cmykColor.magenta = 0;
    cmykColor.yellow = 0;
    cmykColor.black = 100;

    // Define the offset effect
    var xmlstringOffsetPath = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst value I jntp 0 "/></LiveEffect>'; // 0 = round, 1=bevel, 2=miter
    xmlstringOffsetPath = xmlstringOffsetPath.replace("value", offset);

    idoc.selection = null;

    var dup = sel.duplicate(sel, ElementPlacement.PLACEAFTER);
    try {
        dup.filled = false;
    } catch (e) {}

    // Select new path before applying effects
    dup.selected = true;

    // Fill color
    dup.filled = true;
    dup.fillColor = cmykColor;

    // Apply offset path Effect
    dup.applyEffect(xmlstringOffsetPath);

    // Move to layer
    var layerName = 'Outline';
    var _layer = idoc.layers.getByName(layerName);
    dup.move(_layer, ElementPlacement.PLACEATEND);

    // expand appearance
    app.executeMenuCommand('expandStyle');
    app.executeMenuCommand("group");
    app.executeMenuCommand("Live Pathfinder Add");
    app.executeMenuCommand("expandStyle");

    // Redraw
    app.redraw();

    app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
}

main();