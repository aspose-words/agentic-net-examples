using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportToPlainTextExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add several paragraphs using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Write("Third paragraph without a line break.");

        // Set up TxtSaveOptions to define a custom paragraph break.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.ParagraphBreak = " <END>\r\n";

        // Paths for the output files.
        string docxPath = "ExportExample.docx";
        string txtPath = "ExportExample.txt";

        // Save the original document as DOCX (optional, just for reference).
        doc.Save(docxPath, SaveFormat.Docx);

        // Export the document to plain‑text using the custom paragraph break.
        doc.Save(txtPath, txtOptions);

        // Load the exported plain‑text file to verify its contents.
        PlainTextDocument plain = new PlainTextDocument(txtPath);
        Console.WriteLine("Plain‑text content:");
        Console.WriteLine(plain.Text);
    }
}
