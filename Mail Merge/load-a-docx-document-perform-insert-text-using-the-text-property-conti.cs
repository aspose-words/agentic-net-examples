using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the desired text at the current cursor position.
        builder.Write("This is the inserted text.");

        // Save the modified document as a PNG image (first page rendered).
        doc.Save("OutputImage.png", SaveFormat.Png);
    }
}
