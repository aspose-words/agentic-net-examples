using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.mht");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello World!");
        builder.Font.Size = 24;
        builder.Writeln("This document will be converted to MHTML with embedded CSS.");
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Configure save options for MHTML.
        //    - Use SaveFormat.Mhtml.
        //    - Embed CSS directly into the HTML part (CssStyleSheetType.Embedded).
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            CssStyleSheetType = CssStyleSheetType.Embedded,
            // Ensure resources are referenced by file name (default) – CID URLs are not required.
            ExportCidUrlsForMhtmlResources = false
        };

        // -----------------------------------------------------------------
        // 4. Save the document as MHTML.
        // -----------------------------------------------------------------
        doc.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created and contains data.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("MHTML conversion failed: output file was not created or is empty.");
        }

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine("Conversion completed successfully.");
    }
}
