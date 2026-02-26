using System;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("Input.docm");

        // Turn on automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: configure additional hyphenation settings.
        // Set the hyphenation zone to 0.5 inch (720 twentieths of a point).
        doc.HyphenationOptions.HyphenationZone = 720;
        // Allow at most two consecutive lines to end with hyphens.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

        // Rebuild the layout so that hyphenation is applied.
        doc.UpdatePageLayout();

        // Save the modified document.
        doc.Save("Output.docm");
    }
}
