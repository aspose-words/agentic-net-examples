using System;
using System.IO;
using Aspose.Words;

class SplitDocumentMerger
{
    static void Main()
    {
        // Path to the folder that contains the split document parts (DOCX files).
        string splitPartsFolder = @"C:\Docs\SplitParts";

        // Load the first part to create the base merged document.
        string[] partFiles = Directory.GetFiles(splitPartsFolder, "*.docx");
        if (partFiles.Length == 0)
        {
            Console.WriteLine("No split document parts found.");
            return;
        }

        Document mergedDocument = new Document(partFiles[0]);

        // Append the remaining split parts to the merged document.
        for (int i = 1; i < partFiles.Length; i++)
        {
            Document part = new Document(partFiles[i]);
            mergedDocument.AppendDocument(part, ImportFormatMode.KeepSourceFormatting);
        }

        // Load the additional document that should be merged with the split document.
        string otherDocumentPath = @"C:\Docs\OtherDocument.docx";
        Document otherDocument = new Document(otherDocumentPath);

        // Append the other document to the merged result.
        mergedDocument.AppendDocument(otherDocument, ImportFormatMode.KeepSourceFormatting);

        // Save the final merged document.
        string outputPath = @"C:\Docs\MergedResult.docx";
        mergedDocument.Save(outputPath, SaveFormat.Docx);

        Console.WriteLine($"Merged document saved to: {outputPath}");
    }
}
