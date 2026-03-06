using System;
using Aspose.Words;

class MergeDocuments
{
    static void Main()
    {
        // Load the first document (the part that was split off).
        Document splitPart = new Document("SplitPart.docx");

        // Load the second document to which the split part will be merged.
        Document mainDoc = new Document("MainDocument.docx");

        // Append the split part to the end of the main document.
        // KeepSourceFormatting preserves the original formatting of the appended document.
        mainDoc.AppendDocument(splitPart, ImportFormatMode.KeepSourceFormatting);

        // Save the merged result as a DOCX file.
        mainDoc.Save("MergedResult.docx");
    }
}
