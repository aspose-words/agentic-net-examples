using System;
using Aspose.Words;

namespace AsposeWordsMergeExample
{
    class Program
    {
        static void Main()
        {
            // Load the first part of the split document.
            Document splitDocument = new Document("SplitPart.docx");

            // Load the document that should be merged with the split part.
            Document otherDocument = new Document("OtherDocument.docx");

            // Append the second document to the end of the first one.
            // KeepSourceFormatting preserves the original formatting of the appended document.
            splitDocument.AppendDocument(otherDocument, ImportFormatMode.KeepSourceFormatting);

            // Save the merged result.
            splitDocument.Save("MergedDocument.docx");
        }
    }
}
