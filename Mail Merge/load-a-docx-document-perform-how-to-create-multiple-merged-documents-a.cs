using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the source DOCX files that will be merged.
        string[] sourceFiles = new string[]
        {
            @"C:\Docs\Part1.docx",
            @"C:\Docs\Part2.docx",
            @"C:\Docs\Part3.docx"
        };

        // Load the first document – this will be the base document.
        Document mergedDocument = new Document(sourceFiles[0]);

        // Append the remaining documents to the base document.
        for (int i = 1; i < sourceFiles.Length; i++)
        {
            Document src = new Document(sourceFiles[i]);
            mergedDocument.AppendDocument(src, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged document as a PNG image.
        // The Save method determines the format from the extension or the SaveFormat enum.
        mergedDocument.Save(@"C:\Docs\MergedDocument.png", SaveFormat.Png);
    }
}
