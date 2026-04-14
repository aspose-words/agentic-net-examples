using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables; // Required for Table class

public class HyphenationExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and possible hyphenation.
        builder.PageSetup.PageWidth = 300; // points
        builder.PageSetup.LeftMargin = 20;
        builder.PageSetup.RightMargin = 20;

        // Set the language of the text to English (US) – required for hyphenation.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 360; // 0.25 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Write a paragraph with long words that can be hyphenated.
        builder.Writeln(
            "characteristically characteristically characteristically " +
            "characteristically characteristically characteristically.");

        // Insert a table with a single cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Write the same long text inside the cell.
        builder.Writeln(
            "characteristically characteristically characteristically " +
            "characteristically characteristically characteristically.");

        // Suppress hyphenation for the paragraph inside the cell.
        Paragraph cellParagraph = (Paragraph)builder.CurrentParagraph;
        cellParagraph.ParagraphFormat.SuppressAutoHyphens = true;

        // End the table.
        builder.EndTable();

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationExample.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
