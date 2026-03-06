using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the destination document (the one to which we will append).
        Document dstDoc = new Document("Destination.docx");

        // Load the source document (the one to be appended).
        Document srcDoc = new Document("Source.docx");

        // Append the source document to the end of the destination document.
        // KeepSourceFormatting preserves the original formatting of the source.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting combined document.
        dstDoc.Save("Combined.docx");
    }
}
