using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationExample
{
    public static void Main()
    {
        // Create a folder for all temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationExample");
        Directory.CreateDirectory(workDir);

        // Create a minimal hyphenation dictionary for English (en-US).
        // The OpenOffice hyphenation dictionary format requires:
        //   1) Number of patterns.
        //   2) Pattern lines.
        //   3) Blank line.
        //   4) Number of hyphenation exceptions.
        //   5) Exception lines.
        // An empty dictionary (zero patterns, zero exceptions) is still a valid file.
        string dictPath = Path.Combine(workDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath, "0\r\n\r\n0\r\n");

        // Register the dictionary so that hyphenation can be performed.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Verify that the dictionary is registered.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and enable hyphenation.
        // 300 points ≈ 4.17 inches.
        doc.FirstSection.PageSetup.PageWidth = 300.0;
        doc.FirstSection.PageSetup.PageHeight = 842.0; // A4 height for completeness.

        // Write a paragraph containing a long word that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln(
            "Automatic hyphenation demonstrates how long words such as characteristically " +
            "can be split across lines when the document layout requires it.");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: fine‑tune hyphenation behaviour.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch.
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document to DOCX.
        string outPath = Path.Combine(workDir, "Hyphenated.docx");
        doc.Save(outPath, SaveFormat.Docx);

        // Validate that the output file was created.
        if (!File.Exists(outPath))
            throw new FileNotFoundException("The DOCX file was not created.", outPath);

        // Reload the document to ensure settings persisted.
        Document loadedDoc = new Document(outPath);
        if (!loadedDoc.HyphenationOptions.AutoHyphenation)
            throw new InvalidOperationException("Hyphenation option was not preserved after saving.");

        // All checks passed – the example completed successfully.
        Console.WriteLine("Hyphenation enabled, dictionary registered, and DOCX saved to:");
        Console.WriteLine(outPath);
    }
}
