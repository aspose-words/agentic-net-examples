using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfFilePath = @"C:\Docs\SourceDocument.pdf";

        // Load the PDF document. PdfLoadOptions can be used to control loading behavior.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document pdfDoc = new Document(pdfFilePath, loadOptions);

        // -----------------------------------------------------------------
        // Convert the loaded PDF to a Microsoft Word template (.dot) format.
        // The SaveFormat.Dot enumeration value corresponds to the DOT format.
        // -----------------------------------------------------------------
        string dotOutputPath = @"C:\Docs\ConvertedDocument.dot";
        pdfDoc.Save(dotOutputPath, SaveFormat.Dot);

        // -----------------------------------------------------------------
        // Additional conversions to other common formats (optional).
        // Demonstrates how the same Document instance can be saved in
        // different formats by specifying the appropriate SaveFormat value.
        // -----------------------------------------------------------------
        pdfDoc.Save(@"C:\Docs\ConvertedDocument.docx", SaveFormat.Docx); // Word 2007+ document
        pdfDoc.Save(@"C:\Docs\ConvertedDocument.rtf", SaveFormat.Rtf);   // Rich Text Format
        pdfDoc.Save(@"C:\Docs\ConvertedDocument.html", SaveFormat.Html); // HTML format
    }
}
