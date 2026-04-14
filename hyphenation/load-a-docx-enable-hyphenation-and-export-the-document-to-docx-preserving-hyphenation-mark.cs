using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationExample
{
    public static void Main()
    {
        // Define file names in the current directory.
        const string inputFile = "sample_input.docx";
        const string outputFile = "sample_output.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample document with long text that can be hyphenated.
        // -----------------------------------------------------------------
        var createDoc = new Document();
        var builder = new DocumentBuilder(createDoc);

        // Narrow the page width to increase the chance of line breaks.
        builder.PageSetup.PageWidth = 300; // points (~4.17 inches)

        // Write a paragraph containing long words.
        builder.Font.Size = 12;
        builder.Writeln(
            "Antidisestablishmentarianism is often cited as one of the longest words in the English language. " +
            "Supercalifragilisticexpialidocious is another example that may need hyphenation when the line is short.");

        // Save the source document.
        createDoc.Save(inputFile);

        // -----------------------------------------------------------------
        // Step 2: Load the document, enable automatic hyphenation, and save.
        // -----------------------------------------------------------------
        var doc = new Document(inputFile);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document preserving hyphenation marks.
        doc.Save(outputFile);

        // -----------------------------------------------------------------
        // Step 3: Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"Failed to create the output file '{outputFile}'.");

        // Optional: Inform the user (no interactive input required).
        Console.WriteLine($"Hyphenated document saved to '{outputFile}'.");
    }
}
