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
        builder.Writeln("First section content.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Second section content.");

        // Prepare an output directory.
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
            // Attempt to save the document with splitting enabled.
            // This operation is not supported for MHTML and should throw an exception.
            doc.Save(outputPath, saveOptions);
            Console.WriteLine("Document saved successfully (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected: an exception indicating that splitting is not supported for MHTML.
            Console.WriteLine($"Caught expected exception: {ex.GetType().Name} - {ex.Message}");
        }
    }
}
