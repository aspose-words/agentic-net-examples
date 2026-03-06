using System;
using Aspose.Words;

class ConvertDocx
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Convert to PDF.
        string pdfPath = @"C:\Docs\ConvertedDocument.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Convert to HTML.
        string htmlPath = @"C:\Docs\ConvertedDocument.html";
        doc.Save(htmlPath, SaveFormat.Html);

        // Convert to plain text.
        string txtPath = @"C:\Docs\ConvertedDocument.txt";
        doc.Save(txtPath, SaveFormat.Text);
    }
}
