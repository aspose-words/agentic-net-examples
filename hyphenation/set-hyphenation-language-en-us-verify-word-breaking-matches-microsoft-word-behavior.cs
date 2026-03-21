using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Create a temporary hyphenation dictionary file if it does not exist.
        // The file can be empty; Aspose.Words only requires the file to be present.
        string tempDir = Path.GetTempPath();
        string dictionaryPath = Path.Combine(tempDir, "hyph_en_US.dic");

        if (!File.Exists(dictionaryPath))
        {
            // Write a minimal (empty) dictionary file.
            File.WriteAllText(dictionaryPath, string.Empty);
        }

        // Register the English (US) hyphenation dictionary.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Verify that the dictionary has been registered successfully.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the en‑US hyphenation dictionary.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 0;
        doc.HyphenationOptions.HyphenationZone = 360;

        // Write a paragraph that contains a long word which will be hyphenated.
        builder.Font.Size = 24;
        builder.Writeln(
            "The word characteristically demonstrates how automatic hyphenation works when the line width is constrained.");

        // Force a narrow page width to make hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 300;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Update the layout so that hyphenation is applied.
        doc.UpdatePageLayout();

        // Save the document to the temporary directory.
        string outputPath = Path.Combine(tempDir, "HyphenatedDocument.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
