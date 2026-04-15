using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure the font to be bold and underlined.
        builder.Font.Bold = true;
        builder.Font.Underline = Underline.Single;

        // Apply a heading style (optional, but typical for header text).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Insert the formatted header text.
        builder.Writeln("Formatted Header");

        // Ensure the output directory exists.
        string outputDir = "Artifacts";
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "HeaderBoldUnderline.docx");
        doc.Save(outputPath);
    }
}
