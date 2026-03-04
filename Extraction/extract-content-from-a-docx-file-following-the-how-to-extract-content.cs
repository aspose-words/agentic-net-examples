using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document sourceDoc = new Document("Source.docx");

        // Extract plain‑text representation of the document.
        PlainTextDocument plainText = new PlainTextDocument("Source.docx");
        string text = plainText.Text;

        // Create a new blank document and write the extracted text into it.
        Document outputDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(outputDoc);
        builder.Writeln(text);

        // Save the result as a plain‑text file.
        outputDoc.Save("ExtractedText.txt", SaveFormat.Text);
    }
}
