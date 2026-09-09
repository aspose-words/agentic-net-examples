using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary input DOCX and the resulting MHTML file.
        const string inputPath = "sample.docx";
        const string outputPath = "sample.mht";

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document with some text and a shape.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");

        // Insert a rectangle shape so the output contains an image resource.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Save the document as DOCX (required by the task's bootstrap rule).
        doc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX file that we just created.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Configure save options for MHTML with embedded images and fonts.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Export font resources and embed them as Base64 within the MHTML package.
            ExportFontResources = true,
            ExportFontsAsBase64 = true
            // Images are embedded by default when saving to MHTML, so no extra setting is required.
        };

        // -----------------------------------------------------------------
        // 4. Save the document as MHTML.
        // -----------------------------------------------------------------
        loaded.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the MHTML file was created and is not empty.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("MHTML conversion failed: output file was not created or is empty.");
        }

        Console.WriteLine($"Document successfully converted to MHTML: {outputPath}");
    }
}
