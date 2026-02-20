using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Themes;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // -------------------- Font Settings --------------------
        // Initialize FontSettings and assign it to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Specify a folder that contains custom TrueType fonts.
        // The second parameter (false) indicates that subfolders are not searched.
        string fontsFolder = @"C:\MyFonts";
        fontSettings.SetFontsFolder(fontsFolder, false);

        // Build an automatic fallback scheme based on the fonts available in the folder.
        fontSettings.FallbackSettings.BuildAutomatic();

        // Configure a default substitution: if a requested font is missing, use Arial.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // -------------------- Theme Fonts --------------------
        // Access the document's theme and set major/minor fonts for Latin script.
        Theme theme = doc.Theme;
        theme.MajorFonts.Latin = "Times New Roman";
        theme.MinorFonts.Latin = "Calibri";

        // -------------------- Write Sample Text --------------------
        // Use DocumentBuilder to add text that references a missing font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Missing Font"; // Triggers fallback/substitution.
        builder.Writeln("This paragraph uses a missing font and will be rendered with the fallback or substituted font.");

        // -------------------- Save Document --------------------
        // Save the document as DOCX.
        doc.Save(@"C:\Output\ExampleFontSettings.docx");
    }
}
