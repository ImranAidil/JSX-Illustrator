#target illustrator

var win = new Window("dialog", "Color Grid Generator");
win.orientation = "column";
win.alignChildren = "left";

// Input fields
win.add("statictext", undefined, "Number of columns:");
var columnInput = win.add("edittext", undefined, "5");
columnInput.characters = 5;

win.add("statictext", undefined, "Number of rows:");
var rowInput = win.add("edittext", undefined, "5");
rowInput.characters = 5;

win.add("statictext", undefined, "Spacing between objects:");
var spacingInput = win.add("edittext", undefined, "10");
spacingInput.characters = 5;

// Color variation
var colorVariationGroup = win.add("group");
colorVariationGroup.orientation = "row";
var colorVariationCheckbox = colorVariationGroup.add("checkbox", undefined, "Apply color variation");
var colorChannelDropdown = colorVariationGroup.add("dropdownlist", undefined, ["C", "M", "Y", "K"]);
colorChannelDropdown.selection = 0;
colorChannelDropdown.enabled = false;

colorVariationCheckbox.onClick = function() {
    colorChannelDropdown.enabled = colorVariationCheckbox.value;
}

// Buttons
var buttonGroup = win.add("group");
buttonGroup.orientation = "row";
var okButton = buttonGroup.add("button", undefined, "OK");
var cancelButton = buttonGroup.add("button", undefined, "Cancel");

okButton.onClick = function() {
    win.close();
    generateGrid();
}

cancelButton.onClick = function() {
    win.close();
}

win.show();

function generateGrid() {
    try {
        if (app.activeDocument && app.activeDocument.selection.length > 0) {
            var activeObject = app.activeDocument.selection[0];
            var fillColor = activeObject.fillColor;

            var horizontalCount = parseInt(columnInput.text, 10);
            var verticalCount = parseInt(rowInput.text, 10);
            var spacing = parseFloat(spacingInput.text);
            var colorVariation = colorVariationCheckbox.value;
            var colorChannel = colorChannelDropdown.selection.text;

            if (isNaN(horizontalCount) || isNaN(verticalCount) || isNaN(spacing) || 
                horizontalCount <= 0 || verticalCount <= 0 || spacing < 0) {
                alert("Please enter valid positive numbers for rows, columns, and spacing.");
                return;
            }

            var originalX = activeObject.left;
            var originalY = activeObject.top;

            // Convert to CMYK if not already
            var cmykFillColor = convertToCMYK(fillColor);

            for (var i = 0; i < verticalCount; i++) {
                for (var j = 0; j < horizontalCount; j++) {
                    var newObject = activeObject.duplicate();
                    newObject.left = originalX + (j * (activeObject.width + spacing));
                    newObject.top = originalY - (i * (activeObject.height + spacing));

                    var newColor = new CMYKColor();
                    if (colorVariation) {
                        var variationValue = (i + j) / (verticalCount + horizontalCount - 2) * 100;
                        newColor.cyan = (colorChannel === "C") ? variationValue : cmykFillColor.cyan;
                        newColor.magenta = (colorChannel === "M") ? variationValue : cmykFillColor.magenta;
                        newColor.yellow = (colorChannel === "Y") ? variationValue : cmykFillColor.yellow;
                        newColor.black = (colorChannel === "K") ? variationValue : cmykFillColor.black;
                    } else {
                        newColor = cmykFillColor;
                    }
                    newObject.fillColor = newColor;

                    // Add CMYK value text
                    var textFrame = app.activeDocument.textFrames.add();
                    textFrame.contents = "C:" + Math.round(newColor.cyan) + 
                                         " M:" + Math.round(newColor.magenta) + 
                                         " Y:" + Math.round(newColor.yellow) + 
                                         " K:" + Math.round(newColor.black);
                    
                    // Set initial text frame position at the bottom of the object
                    var docCoords = newObject.position;
                    textFrame.left = docCoords[0] + newObject.width * 0.1; // 10% margin from left
                    textFrame.top = docCoords[1] - newObject.height * 0.9; // 10% margin from bottom
                    
                    // Style the text
                    var textRange = textFrame.textRange;
                    textRange.characterAttributes.fillColor = getContrastColor(newColor);
                    textRange.characterAttributes.size = 12; // Set a default font size

                    // Left align the text
                    textRange.paragraphAttributes.justification = Justification.LEFT;

                    // Fit text to frame
                    fitTextToObject(textFrame, newObject);

                    // Update progress (for large grids)
                    if ((i * horizontalCount + j + 1) % 100 === 0) {
                        $.writeln("Progress: " + Math.round(((i * horizontalCount + j + 1) / (horizontalCount * verticalCount)) * 100) + "%");
                    }
                }
            }
            alert("Grid creation complete!");
        } else {
            alert("Please select an object to duplicate.");
        }
    } catch (e) {
        alert("An error occurred: " + e.message);
    }
}

function fitTextToObject(textFrame, object) {
    var maxWidth = object.width * 0.9; // 90% of the object width
    var maxHeight = object.height * 0.2; // 20% of the object height
    var fontSize = 12; // Start with 12pt font size

    textFrame.textRange.characterAttributes.size = fontSize;

    while (textFrame.width > maxWidth || textFrame.height > maxHeight) {
        fontSize -= 0.5;
        if (fontSize < 4) break; // Minimum font size
        textFrame.textRange.characterAttributes.size = fontSize;
    }

    // Position the text at the bottom of the object
    textFrame.top = object.top - object.height + textFrame.height;
}

function convertToCMYK(color) {
    if (color.typename === "CMYKColor") {
        return color;
    } else if (color.typename === "RGBColor") {
        var cmykColor = new CMYKColor();
        var rgb = [color.red, color.green, color.blue];
        var cmyk = rgb2cmyk(rgb);
        cmykColor.cyan = cmyk[0];
        cmykColor.magenta = cmyk[1];
        cmykColor.yellow = cmyk[2];
        cmykColor.black = cmyk[3];
        return cmykColor;
    } else if (color.typename === "GrayColor") {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = color.gray;
        return cmykColor;
    } else {
        // For other color types, return a default CMYK color
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = 100;
        return cmykColor;
    }
}

function rgb2cmyk(rgb) {
    var r = rgb[0] / 255;
    var g = rgb[1] / 255;
    var b = rgb[2] / 255;
    
    var k = 1 - Math.max(r, g, b); // Black key
    var c = (1 - r - k) / (1 - k); // Cyan
    var m = (1 - g - k) / (1 - k); // Magenta
    var y = (1 - b - k) / (1 - k); // Yellow
    
    return [c * 100, m * 100, y * 100, k * 100];
}

function getContrastColor(cmykColor) {
    // Calculate perceived brightness
    var brightness = (1 - cmykColor.cyan/100) * 0.3 + 
                     (1 - cmykColor.magenta/100) * 0.59 + 
                     (1 - cmykColor.yellow/100) * 0.11;

    // Factor in the black component
    brightness *= (1 - cmykColor.black/100);

    var contrastColor = new CMYKColor();
    if (brightness < 0.5) {
        contrastColor.cyan = 0;
        contrastColor.magenta = 0;
        contrastColor.yellow = 0;
        contrastColor.black = 0; // White text for dark backgrounds
    } else {
        contrastColor.cyan = 0;
        contrastColor.magenta = 0;
        contrastColor.yellow = 0;
        contrastColor.black = 100; // Black text for light backgrounds
    }
    return contrastColor;
}