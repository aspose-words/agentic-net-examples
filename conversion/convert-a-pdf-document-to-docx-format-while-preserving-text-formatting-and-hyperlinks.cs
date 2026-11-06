using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file with text formatting and a hyperlink.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Font.Size = 12;
        builder.Writeln("This is a sample PDF document.");
        builder.Font.Size = 14;
        builder.Font.Bold = true;
        builder.Writeln("It contains formatted text and a hyperlink:");
        builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);
        builder.Writeln();
        sourceDoc.Save("input.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("input.pdf");

        // Convert and save as DOCX, preserving formatting and hyperlinks.
        pdfDoc.Save("output.docx", SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists("output.docx"))
        {
            throw new InvalidOperationException("The DOCX output file was not created.");
        }
    }
}
