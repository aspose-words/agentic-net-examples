using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First section");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Second section");

        // Prepare an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "SplitMhtml.mhtml");

        // Configure HtmlSaveOptions for MHTML with a split criterion.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        try
        {
            // Attempt to save; this should fail because splitting is not supported for MHTML.
            doc.Save(outputPath, saveOptions);
            // If no exception is thrown, indicate unexpected behavior.
            Console.WriteLine("Document saved without exception – this is unexpected for MHTML splitting.");
        }
        catch (Exception ex)
        {
            // Expected path: an exception is thrown.
            Console.WriteLine($"Caught expected exception: {ex.GetType().Name} - {ex.Message}");
        }
    }
}
