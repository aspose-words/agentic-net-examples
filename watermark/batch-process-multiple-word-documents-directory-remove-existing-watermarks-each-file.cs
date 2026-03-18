using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Folder that contains the Word documents to process
        string inputFolder = @"C:\Docs\Input";

        // Get all .docx files in the folder (adjust the pattern if other formats are needed)
        string[] files = Directory.GetFiles(inputFolder, "*.docx");

        foreach (string filePath in files)
        {
            // Load the document from file (uses Document(string) constructor rule)
            Document doc = new Document(filePath);

            // If the document has any watermark (text, image, or other), remove it
            if (doc.Watermark.Type != WatermarkType.None)
                doc.Watermark.Remove(); // Uses Watermark.Remove() method rule

            // Save the modified document back to the same file (uses Document.Save(string) rule)
            doc.Save(filePath);
        }
    }
}
