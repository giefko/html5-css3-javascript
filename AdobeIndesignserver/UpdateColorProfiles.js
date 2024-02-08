var app;
try {
    app = new ActiveXObject("InDesignServer.Application");
} catch (e) {
    WScript.Echo("Unable to create InDesign Server Application object.");
    WScript.Quit();
}

// Enable color management
app.colorSettings.enableColorManagement = true;

// Set new color profile settings
app.colorSettings.workingSpaceRGB = "New RGB Profile";
app.colorSettings.workingSpaceCMYK = "New CMYK Profile";

// Log updated color profile settings
var currentRGBProfile = app.colorSettings.workingSpaceRGB;
var currentCMYKProfile = app.colorSettings.workingSpaceCMYK;

WScript.Echo("Updated RGB Color Profile: " + currentRGBProfile);
WScript.Echo("Updated CMYK Color Profile: " + currentCMYKProfile);
