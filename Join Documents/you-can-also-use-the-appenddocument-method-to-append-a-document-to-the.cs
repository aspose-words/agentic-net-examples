using System;
using Aspose.Words;

namespace AppendDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Load the destination document (or create a new blank one).
            Document dstDoc = new Document("Destination.docx");

            // Load the source document that will be appended.
            Document srcDoc = new Document("Source.docx");

            // Append the source document to the end of the destination document.
            // KeepSourceFormatting preserves the original formatting of the source.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document.
            dstDoc.Save("CombinedResult.docx");
        }
    }
}
