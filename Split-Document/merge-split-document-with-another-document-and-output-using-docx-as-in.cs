using System;
using Aspose.Words;
using Aspose.Words.Saving;

class MergeDocuments
{
    static void Main()
    {
        // Load the first document (the split part) from a DOCX file.
        Document splitDoc = new Document("SplitPart.docx");

        // Load the second document (the document to merge with) from a DOCX file.
        Document otherDoc = new Document("OtherDocument.docx");

        // Append the split document to the end of the other document.
        // KeepSourceFormatting preserves the original formatting of the appended content.
        otherDoc.AppendDocument(splitDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged result as a new DOCX file.
        otherDoc.Save("MergedResult.docx");
    }
}
