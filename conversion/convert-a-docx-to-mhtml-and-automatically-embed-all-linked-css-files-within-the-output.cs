using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.mht");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document to act as the source.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be converted to MHTML.");
        builder.Writeln("The conversion will embed all CSS into the resulting file.");
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Configure save options for MHTML with embedded CSS.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Embed CSS directly into the MHTML file instead of linking to external files.
            CssStyleSheetType = CssStyleSheetType.Embedded,

            // Optional: use CID URLs for resources; not required for CSS embedding but can improve compatibility.
            ExportCidUrlsForMhtmlResources = false
        };

        // -----------------------------------------------------------------
        // 4. Save the document as MHTML.
        // -----------------------------------------------------------------
        doc.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The MHTML output file was not created.");

        // The example finishes here; no interactive prompts are used.
    }
}
