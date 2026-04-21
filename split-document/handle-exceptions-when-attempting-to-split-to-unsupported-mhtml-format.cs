using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitMhtmlExample
{
    public static void Main()
    {
        // Define a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First section.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Second section.");

        // Prepare save options for MHTML format and request splitting by section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Attempt to save; MHTML does not support splitting, so an exception is expected.
        try
        {
            string outPath = Path.Combine(artifactsDir, "SplitMhtml.mhtml");
            doc.Save(outPath, saveOptions);
            Console.WriteLine("Document saved successfully (unexpected).");
        }
        catch (Exception ex)
        {
            // Handle the expected exception gracefully.
            Console.WriteLine($"Exception caught while splitting to MHTML: {ex.Message}");
        }
    }
}
