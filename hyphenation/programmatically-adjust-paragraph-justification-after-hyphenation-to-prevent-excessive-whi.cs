using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        // 300 points ≈ 4.17 inches.
        builder.PageSetup.PageWidth = 300.0;

        // Sample long text that will require hyphenation.
        string longText = "Aspose.Words provides powerful document processing capabilities, allowing developers to create, edit, convert, and render Word documents programmatically. " +
                          "When paragraphs are justified, excessive white‑space gaps may appear if hyphenation is not applied correctly.";

        // Write the text into the document.
        builder.Font.Size = 12;
        builder.Writeln(longText);

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Reduce the hyphenation zone so hyphenation can occur closer to the right margin.
        // The property expects a positive value (in 1/20 point). Use a small non‑zero value.
        doc.HyphenationOptions.HyphenationZone = 10; // 0.5 point
        // Limit consecutive hyphenated lines to avoid long runs of hyphens.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

        // Adjust justification mode to compress spacing after hyphenation,
        // which reduces large gaps in justified paragraphs.
        doc.JustificationMode = JustificationMode.Compress;

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationAdjusted.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
