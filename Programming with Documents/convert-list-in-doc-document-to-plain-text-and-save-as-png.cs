using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the original DOC document.
        Document sourceDoc = new Document("Input.doc");

        // Convert the document content to plain text.
        string plainText = sourceDoc.ToString(SaveFormat.Text).Trim();

        // Create a new blank document and write the extracted text into it.
        Document textDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(textDoc);
        builder.Writeln(plainText);

        // Save the new document as a PNG image (renders the first page).
        textDoc.Save("Output.png", SaveFormat.Png);
    }
}
