/*
Code for Import https://scriptui.joonas.me — (Triple click to select): 
{"activeId":2,"items":{"item-0":{"id":0,"type":"Dialog","parentId":false,"style":{"enabled":true,"varName":null,"windowType":"Dialog","creationProps":{"su1PanelCoordinates":false,"maximizeButton":false,"minimizeButton":false,"independent":false,"closeButton":true,"borderless":false,"resizeable":false},"text":"Fredes Tools","preferredSize":[200,200],"margins":16,"orientation":"column","spacing":10,"alignChildren":["center","top"]}},"item-1":{"id":1,"type":"Button","parentId":0,"style":{"enabled":true,"varName":"centerButton","text":"Center Anchor Point in Comp","justify":"center","preferredSize":[0,0],"alignment":"fill","helpTip":"Centers the anchor points in the middle of the composition, without moving the layer."}},"item-2":{"id":2,"type":"StaticText","parentId":0,"style":{"enabled":false,"varName":null,"creationProps":{"truncate":"none","multiline":false,"scrolling":false},"softWrap":false,"text":"Made by <b>Frede Knudsen</b>","justify":"center","preferredSize":[0,0],"alignment":"fill","helpTip":null}}},"order":[0,1,2],"settings":{"importJSON":true,"indentSize":false,"cepExport":false,"includeCSSJS":true,"showDialog":true,"functionWrapper":false,"afterEffectsDockable":false,"itemReferenceList":"None"}}
*/

// DIALOG
// ======


var dialog = new Window("palette", "Fredes Tools", undefined, { resizeable: true, closeButton: true });
dialog.text = "Fredes Tools";
dialog.preferredSize.width = 200;
dialog.preferredSize.height = 200;
dialog.orientation = "column";
dialog.alignChildren = ["center", "top"];
dialog.spacing = 10;
dialog.margins = 16;

var centerButton = dialog.add("button", undefined, undefined, { name: "centerButton" });
centerButton.helpTip = "Centers the anchor points in the middle of the composition, without moving the layer.";
centerButton.text = "Center Anchor Point in Comp";
centerButton.alignment = ["fill", "top"];
centerButton.onClick = centerAnchorInComp;

var byline = dialog.add("statictext", undefined, undefined, { name: "byline" });
byline.text = "Made by Frede Knudsen";
byline.justify = "center";
byline.alignment = ["fill", "top"];

function centerAnchorInComp() {
    // Get active composition
    var activeComp = getActiveComp();

    // Get all selected layers
    var activeLayers = getActiveLayers(activeComp);

    console.log(activeLayers);

    // Get active comp size
    //var compSize = [activeComp.width, activeComp.height];
    //var compCenter = [compSize[0] * 0.5, compSize[1] * 0.5];

    // For each selected layer
    for (var i = 0; i < activeLayers.length; i++) {
        var layer = activeLayers[i];
        var anchorPoint = layer.anchorPoint;
        var transformPosition = layer.transformPosition;

        alert("Anchor point is: " + anchorPoint.value);

        var anchorOffset = [anchorPoint.value[0] - transformPosition.value[0], anchorPoint.value[1] - transformPosition.value[1]];

        alert("Offset is: " + anchorOffset);
    }
}



function getActiveComp() {
    app.activeViewer.setActive();
    return app.project.activeItem;
}

function getActiveLayers(comp) {
    return comp.selectedLayers;
}





dialog.show();