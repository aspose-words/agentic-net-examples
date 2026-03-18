using System;
using Aspose.Words;

class ConvertPdfToXps
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\document.pdf";

        // Path where the XPS file will be saved.
        string xpsPath = @"C:\Output\document.xps";

        // Load the PDF document.
        Document doc = new Document(pdfPath);

        // Save the document in XPS format for Windows XPS Viewer compatibility.
        doc.Save(xpsPath, SaveFormat.Xps);
    }
}
