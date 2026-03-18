using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Root directory that contains multiple subfolders with DOCX files.
        string rootFolder = @"C:\DocsRoot";

        // Create an empty master document.
        Document masterDoc = new Document();

        // Retrieve all DOCX files from the root folder and its subfolders.
        string[] docxFiles = Directory.GetFiles(rootFolder, "*.docx", SearchOption.AllDirectories);

        // Append each document to the master using destination styles.
        foreach (string filePath in docxFiles)
        {
            // Load the source document from file.
            Document srcDoc = new Document(filePath);

            // Append the source document to the master document.
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
        }

        // Export the combined document to PDF.
        masterDoc.Save(@"C:\Output\MasterDocument.pdf");
    }
}
