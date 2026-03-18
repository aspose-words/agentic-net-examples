using System;
using System.IO;
using Aspose.Words;

class InsertExternalHtml
{
    static void Main()
    {
        // Path to the HTML snippet file.
        string htmlFilePath = @"C:\Temp\snippet.html";

        // Path where the resulting Word document will be saved.
        string outputDocPath = @"C:\Temp\Result.docx";

        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Read the HTML content from the external file.
        string htmlContent = File.ReadAllText(htmlFilePath);

        // Insert the HTML content into the document at the current cursor position.
        builder.InsertHtml(htmlContent);

        // Save the document to the specified location.
        doc.Save(outputDocPath);
    }
}
