var app;
try {
    app = new ActiveXObject("InDesignServer.Application");
} catch (e) {
    WScript.Echo("Unable to create InDesign Server Application object.");
    WScript.Quit();
}

// Enable color management
app.colorSettings.enableColorManagement = true;

// Log current color profile settings
var currentRGBProfile = app.colorSettings.workingSpaceRGB;
var currentCMYKProfile = app.colorSettings.workingSpaceCMYK;

WScript.Echo("Current RGB Color Profile: " + currentRGBProfile);
WScript.Echo("Current CMYK Color Profile: " + currentCMYKProfile);

// Close the instance
app = null;
