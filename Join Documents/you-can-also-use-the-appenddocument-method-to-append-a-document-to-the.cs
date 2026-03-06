using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a blank destination document.
        Document dstDoc = new Document();

        // Add initial content to the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.Writeln("Destination document text.");

        // Load the source document from a file.
        Document srcDoc = new Document("Source.docx");

        // Append the source document to the end of the destination document,
        // preserving the source document's formatting.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document to disk.
        dstDoc.Save("Combined.docx");
    }
}
