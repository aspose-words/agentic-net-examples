using System;
using Aspose.Words;

class MergeSplitDocument
{
    static void Main()
    {
        // Path to the split document (part of the original document)
        string splitDocPath = @"C:\Docs\SplitPart.docx";

        // Path to the document that will be merged with the split document
        string otherDocPath = @"C:\Docs\Other.docx";

        // Path where the merged result will be saved
        string outputPath = @"C:\Docs\MergedResult.docx";

        // Load the split document
        Document splitDoc = new Document(splitDocPath);

        // Load the other document
        Document otherDoc = new Document(otherDocPath);

        // Append the other document to the end of the split document.
        // KeepSourceFormatting preserves the original formatting of the appended document.
        splitDoc.AppendDocument(otherDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document
        splitDoc.Save(outputPath);
    }
}
