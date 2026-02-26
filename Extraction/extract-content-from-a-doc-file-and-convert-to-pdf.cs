using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToPdfConverter
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputFile = @"C:\Input\SampleDocument.doc";

        // Path to the destination PDF file.
        string outputFile = @"C:\Output\SampleDocument.pdf";

        // Load the DOC file into an Aspose.Words Document object.
        Document doc = new Document(inputFile);

        // Convert and save the document as PDF.
        // The SaveFormat.Pdf enum explicitly specifies the PDF format.
        doc.Save(outputFile, SaveFormat.Pdf);
    }
}
