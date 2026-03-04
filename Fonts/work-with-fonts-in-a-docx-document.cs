using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class FontDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and format fonts.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set basic font properties.
        builder.Font.Name = "Courier New";
        builder.Font.Size = 24;               // 24 points.
        builder.Font.Bold = true;             // Bold text.
        builder.Font.Italic = true;           // Italic text.
        builder.Font.Color = Color.DarkBlue;  // Font color.
        builder.Font.Underline = Underline.Double; // Double underline.
        builder.Font.HighlightColor = Color.Yellow; // Highlight.

        // Add a border around the text.
        builder.Font.Border.Color = Color.Green;
        builder.Font.Border.LineWidth = 1.5;
        builder.Font.Border.LineStyle = LineStyle.DashDotStroker;

        // Write the formatted text.
        builder.Writeln("Formatted text with custom font settings.");

        // -----------------------------------------------------------------
        // Demonstrate using FontSettings to add a custom font folder.
        // -----------------------------------------------------------------
        // Assume there is a folder "CustomFonts" containing TrueType fonts.
        string customFontsFolder = Path.Combine(Environment.CurrentDirectory, "CustomFonts");

        if (Directory.Exists(customFontsFolder))
        {
            // Add the folder as an additional font source.
            FontSettings.DefaultInstance.SetFontsFolder(customFontsFolder, recursive: true);
        }

        // -----------------------------------------------------------------
        // Demonstrate accessing FontInfos to embed fonts when saving.
        // -----------------------------------------------------------------
        // Enable embedding of all used TrueType fonts.
        FontInfoCollection fontInfos = doc.FontInfos;
        fontInfos.EmbedTrueTypeFonts = true;
        fontInfos.EmbedSystemFonts = true;
        fontInfos.SaveSubsetFonts = true; // Save only the glyphs that are used.

        // Save the document to a DOCX file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FormattedDocument.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
