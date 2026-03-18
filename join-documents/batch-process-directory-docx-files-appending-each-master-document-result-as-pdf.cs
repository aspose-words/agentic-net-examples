using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Folder that contains the DOCX files to be merged
        string sourceFolder = @"C:\Docs\Source";

        // Path for the final merged PDF document
        string outputPdfPath = @"C:\Docs\Merged.pdf";

        // Create an empty master document
        Document masterDoc = new Document();

        // Retrieve all DOCX files in the specified folder (non‑recursive)
        string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx");

        foreach (string filePath in docxFiles)
        {
            // Load each source document
            Document srcDoc = new Document(filePath);

            // Append the source document to the master while keeping its original formatting
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the combined document as PDF; format is inferred from the .pdf extension
        masterDoc.Save(outputPdfPath);
    }
}
