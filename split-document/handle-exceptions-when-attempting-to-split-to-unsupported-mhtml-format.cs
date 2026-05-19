using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        if (!Directory.Exists(artifactsDir))
            Directory.CreateDirectory(artifactsDir);

        // Create a simple document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First section content.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Second section content.");

        // Attempt to split the document while saving to MHTML, which is unsupported.
        string outputPath = Path.Combine(artifactsDir, "SplitMhtml.mhtml");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak; // This should trigger an exception.

        try
        {
            doc.Save(outputPath, saveOptions);
            Console.WriteLine("Document saved successfully (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected path: splitting to MHTML is not supported.
            Console.WriteLine($"Expected exception caught: {ex.Message}");
        }

        // Verify that the file was not created due to the exception.
        if (!File.Exists(outputPath))
        {
            Console.WriteLine("No output file was created, as expected for unsupported split operation.");
        }
        else
        {
            Console.WriteLine("Output file exists (unexpected).");
        }
    }
}
