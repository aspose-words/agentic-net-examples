using System;
using Aspose.Words;
using Aspose.Words.Saving;

class MergeDocumentsToPng
{
    static void Main()
    {
        // Paths to the source DOCX files to be merged.
        string[] sourceFiles = new string[]
        {
            @"C:\Docs\Part1.docx",
            @"C:\Docs\Part2.docx",
            @"C:\Docs\Part3.docx"
        };

        // Load the first document – this will be the destination document.
        Document mergedDoc = new Document(sourceFiles[0]);

        // Append the remaining documents to the destination document.
        for (int i = 1; i < sourceFiles.Length; i++)
        {
            Document srcDoc = new Document(sourceFiles[i]);
            mergedDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged document as PNG. Each page will be rendered to a separate PNG file.
        mergedDoc.Save(@"C:\Docs\MergedDocument.png", SaveFormat.Png);
    }
}
