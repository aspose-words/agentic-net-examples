using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextPng
{
    static void Main()
    {
        // Path to the source DOC file containing the list.
        string sourceDocPath = "input.doc";

        // Path where the resulting PNG image will be saved.
        string outputPngPath = "output.png";

        // Load the original document.
        Document originalDoc = new Document(sourceDocPath);

        // Extract plain‑text representation of the document (including the list).
        PlainTextDocument plainTextDoc = new PlainTextDocument(sourceDocPath);
        string plainText = plainTextDoc.Text;

        // Create a new blank document and write the extracted plain text into it.
        Document textDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(textDoc);
        builder.Writeln(plainText);

        // Configure image save options to render the document as a PNG.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        // Save the plain‑text document as a PNG image.
        textDoc.Save(outputPngPath, pngOptions);
    }
}
