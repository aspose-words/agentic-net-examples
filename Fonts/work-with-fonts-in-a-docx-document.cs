using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Themes;
using System.Drawing;

class FontDemo
{
    static void Main()
    {
        // Define a folder where the resulting document will be saved.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new empty document.
        Document doc = new Document();

        // DocumentBuilder provides a convenient way to add content and set formatting.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Example 1: Basic font properties using DocumentBuilder.Font.
        // -----------------------------------------------------------------
        builder.Font.Name = "Courier New";          // Font family.
        builder.Font.Size = 24;                     // Font size in points.
        builder.Font.Bold = true;                   // Bold style.
        builder.Font.Color = Color.Blue;            // Font color.
        builder.Font.Underline = Underline.Dash;    // Dash underline.
        builder.Writeln("This line uses Courier New, 24pt, bold, blue, dash underline.");

        // -----------------------------------------------------------------
        // Example 2: Using theme font and theme color.
        // -----------------------------------------------------------------
        builder.Font.ThemeFont = ThemeFont.Major;   // Use the document's major theme font.
        builder.Font.ThemeColor = ThemeColor.Accent5; // Use an accent color from the theme.
        builder.Font.TintAndShade = 0.3;            // Lighten the theme color.
        builder.Writeln("This line uses a major theme font and accent5 color with tint.");

        // -----------------------------------------------------------------
        // Example 3: Change the document's theme fonts globally.
        // -----------------------------------------------------------------
        // These settings affect any style that references the theme fonts.
        doc.Theme.MajorFonts.Latin = "Arial";
        doc.Theme.MinorFonts.Latin = "Times New Roman";

        // -----------------------------------------------------------------
        // Example 4: Directly modify a Run's Font properties.
        // -----------------------------------------------------------------
        Run run = new Run(doc, "Run with custom font: Verdana, 18pt, italic.");
        run.Font.Name = "Verdana";
        run.Font.Size = 18;
        run.Font.Italic = true;

        // Insert a new paragraph and add the run to it.
        builder.InsertParagraph();
        builder.CurrentParagraph.AppendChild(run);

        // -----------------------------------------------------------------
        // Save the document to the output folder.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(artifactsDir, "FontDemo.docx");
        doc.Save(outputPath);
    }
}
