using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be converted to MHTML.");
        sourceDoc.Save("sample.docx", SaveFormat.Docx);

        // Load the DOCX document we just created.
        Document doc = new Document("sample.docx");

        // Configure save options for MHTML with embedded CSS.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Embed CSS directly into the HTML part of the MHTML file.
            CssStyleSheetType = CssStyleSheetType.Embedded
        };

        // Save the document as MHTML.
        string outputPath = "output.mht";
        doc.Save(outputPath, saveOptions);

        // Validate that the output file was created and contains data.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("MHTML conversion failed: output file was not created or is empty.");
        }

        // Optionally, inform that the conversion succeeded.
        Console.WriteLine($"Document successfully converted to MHTML: {Path.GetFullPath(outputPath)}");
    }
}
