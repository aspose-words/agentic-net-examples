using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationItalianTest
{
    public static void Main()
    {
        // Create a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "HyphenationItalianTest_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);

        // Prepare a minimal Italian hyphenation dictionary file.
        // The OpenOffice hyphenation dictionary format expects the first line to be the number of patterns.
        // An empty pattern list (0) is sufficient for registration validation.
        string dictPath = Path.Combine(workDir, "hyph_it_IT.dic");
        File.WriteAllText(dictPath, "0\n");

        // Ensure the dictionary is not already registered.
        if (Hyphenation.IsDictionaryRegistered("it-IT"))
            Hyphenation.UnregisterDictionary("it-IT");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("it-IT", dictPath);

        // Verify registration succeeded.
        if (!Hyphenation.IsDictionaryRegistered("it-IT"))
            throw new InvalidOperationException("Italian hyphenation dictionary registration failed.");

        // Create a document with Italian text that would require hyphenation on a narrow page.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a relatively large font to increase the chance of line breaking.
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("it-IT").LCID;

        // A long Italian word that can be hyphenated.
        builder.Writeln("anticonstituzionalmente");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Reduce the page width to force wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200; // points (~2.78 inches)

        // Save the document to PDF.
        string pdfPath = Path.Combine(workDir, "ItalianHyphenation.pdf");
        doc.Save(pdfPath);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The PDF output was not created.", pdfPath);

        // Clean up: unregister the dictionary (optional).
        Hyphenation.UnregisterDictionary("it-IT");

        // Indicate success.
        Console.WriteLine("Hyphenation test completed successfully.");
    }
}
