using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a simple document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First section.");

        // Insert a section break that starts a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Second section.");

        // Path for the MHTML file we will try to create.
        string mhtmlPath = Path.Combine(outputDir, "Sample.mhtml");

        // Configure HtmlSaveOptions for MHTML with a split criterion.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        try
        {
            // This save operation should fail because splitting is not supported for MHTML.
            doc.Save(mhtmlPath, saveOptions);
            Console.WriteLine("Document saved successfully (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected: an exception indicating that splitting is not supported.
            Console.WriteLine($"Expected exception caught: {ex.Message}");
        }

        // Verify that the MHTML file was not created.
        if (!File.Exists(mhtmlPath))
        {
            Console.WriteLine("MHTML file was not created, as expected.");
        }
        else
        {
            Console.WriteLine("MHTML file exists (unexpected).");
        }
    }
}
