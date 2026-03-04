using System;
using Aspose.Words;

class SaveDocumentAsPdf
{
    static void Main()
    {
        // Path to the folder that contains the input document.
        // Adjust this path as needed for your environment.
        string dataDir = @"C:\MyDocuments\";

        // Load an existing Word document (e.g., DOCX) from the file system.
        // The Document constructor automatically detects the format.
        Document doc = new Document(dataDir + "input.docx");

        // Save the loaded document as PDF.
        // The Save method determines the format from the file extension.
        doc.Save(dataDir + "output.pdf");
    }
}
