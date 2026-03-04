using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string docxPath = @"C:\Input\source.docx";

        // Path where the resulting PDF will be saved.
        string pdfPath = @"C:\Output\result.pdf";

        // Load the DOCX document using the Document(string) constructor.
        Document doc = new Document(docxPath);

        // Save the document as PDF. The Save method determines the format from the file extension.
        doc.Save(pdfPath);
    }
}
