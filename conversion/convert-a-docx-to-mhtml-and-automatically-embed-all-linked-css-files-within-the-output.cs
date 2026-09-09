using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Styled Heading");
        string inputPath = "sample.docx";
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document loaded = new Document(inputPath);

        // Configure save options to produce MHTML with embedded CSS.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            CssStyleSheetType = CssStyleSheetType.Inline,
            ExportCidUrlsForMhtmlResources = false
        };

        // Save as MHTML.
        string outputPath = "output.mht";
        loaded.Save(outputPath, saveOptions);

        // Verify that the output file was created and is not empty.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("MHTML conversion failed: output file not created or empty.");
        }
    }
}
