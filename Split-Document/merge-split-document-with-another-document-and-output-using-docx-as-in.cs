using System;
using Aspose.Words;

namespace DocumentMergingExample
{
    class Program
    {
        static void Main()
        {
            // Load the first (split) document from a DOCX file.
            Document splitDoc = new Document("SplitDocument.docx");

            // Load the second document that will be merged with the first one.
            Document otherDoc = new Document("OtherDocument.docx");

            // Append the second document to the first one.
            // KeepSourceFormatting preserves the original formatting of the appended document.
            splitDoc.AppendDocument(otherDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the merged result as a DOCX file.
            splitDoc.Save("MergedOutput.docx");
        }
    }
}
