using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Set up directories for input and output.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple source document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First section.");
        // Insert a section break (new page) to separate sections.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Second section.");

        // Configure HtmlSaveOptions for MHTML format and request splitting by section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        string outputPath = Path.Combine(artifactsDir, "SplitDocument.mhtml");

        try
        {
            // This operation is not supported for MHTML and should throw an exception.
            doc.Save(outputPath, saveOptions);
            Console.WriteLine("Document saved successfully (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected exception handling.
            Console.WriteLine("Caught exception while attempting to split to MHTML:");
            Console.WriteLine(ex.Message);
        }

        // Verify that no output file was created.
        if (!File.Exists(outputPath))
        {
            Console.WriteLine("No output file was created, as expected for unsupported split.");
        }
        else
        {
            Console.WriteLine("Output file was created unexpectedly.");
        }
    }
}
