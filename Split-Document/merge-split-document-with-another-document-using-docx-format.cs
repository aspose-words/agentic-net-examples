using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentMerger
{
    static void Main()
    {
        // Load the first (split) document.
        Document splitDoc = new Document("SplitPart.docx");

        // Load the second document that we want to merge with the split document.
        Document otherDoc = new Document("AdditionalContent.docx");

        // Append the second document to the first one.
        // ImportFormatMode.KeepSourceFormatting preserves the original formatting of the appended document.
        splitDoc.AppendDocument(otherDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged result as a new DOCX file.
        splitDoc.Save("MergedResult.docx", SaveFormat.Docx);
    }
}
