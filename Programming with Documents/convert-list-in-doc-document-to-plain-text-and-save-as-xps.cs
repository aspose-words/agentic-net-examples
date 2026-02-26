using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the original DOC document.
        Document sourceDoc = new Document("Input.doc");

        // Extract the plain‑text representation (including list labels).
        PlainTextDocument plainTextDoc = new PlainTextDocument("Input.doc");
        string plainText = plainTextDoc.Text;

        // Create a new blank document and write the extracted text into it.
        Document textDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(textDoc);
        builder.Writeln(plainText);

        // Save the new document as XPS.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        textDoc.Save("Output.xps", xpsOptions);
    }
}
