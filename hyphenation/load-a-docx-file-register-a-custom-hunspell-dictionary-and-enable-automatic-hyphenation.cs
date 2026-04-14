using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationExample
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationDemo");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document with long text that can be hyphenated.
        // -----------------------------------------------------------------
        string inputPath = Path.Combine(workDir, "Input.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a language that requires hyphenation (English in this case).
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 24;

        // A paragraph with a long word that will be split when hyphenation is active.
        builder.Writeln(
            "This is a demonstration of automatic hyphenation. " +
            "The word hyphenationexampleisverylongshouldbreak demonstrates the effect.");

        // Save the source document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Create a minimal custom Hunspell hyphenation dictionary.
        //    The OpenOffice hyphenation dictionary format expects the first line
        //    to contain the number of patterns, followed by the patterns themselves.
        // -----------------------------------------------------------------
        string dictPath = Path.Combine(workDir, "hyph_en_US.dic");
        using (StreamWriter writer = new StreamWriter(dictPath))
        {
            // One simple pattern that allows a hyphen after "hy".
            writer.WriteLine("1");
            writer.WriteLine("hy1ph");
        }

        // -----------------------------------------------------------------
        // 3. Register the custom dictionary for the "en-US" locale.
        // -----------------------------------------------------------------
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // -----------------------------------------------------------------
        // 4. Enable automatic hyphenation for the document.
        // -----------------------------------------------------------------
        doc.HyphenationOptions.AutoHyphenation = true;

        // -----------------------------------------------------------------
        // 5. Save the document to PDF so that hyphenation can be observed.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "Output.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 6. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF output was not created.");

        // Clean up (optional): uncomment the following line to delete temporary files after execution.
        // Directory.Delete(workDir, true);
    }
}
