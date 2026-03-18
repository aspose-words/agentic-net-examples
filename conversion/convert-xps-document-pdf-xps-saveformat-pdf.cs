using System;
using Aspose.Words;

class XpsToPdfConverter
{
    static void Main()
    {
        // Path to the source XPS document.
        string xpsPath = "input.xps";

        // Path where the resulting PDF will be saved.
        string pdfPath = "output.pdf";

        // Load the XPS document into an Aspose.Words Document object.
        Document doc = new Document(xpsPath);

        // Save the document as PDF using the SaveFormat.Pdf enumeration.
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
