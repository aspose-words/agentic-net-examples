using System;
using Aspose.Words;

class MergeDocuments
{
    static void Main()
    {
        // Load the first (destination) part of the split document.
        Document destination = new Document("Part1.docx");

        // Load the second part that will be merged into the first.
        Document source = new Document("Part2.docx");

        // Append the source document to the end of the destination document.
        // KeepSourceFormatting preserves the original formatting of the source.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document in DOCX format.
        destination.Save("MergedDocument.docx");
    }
}
