using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;

public class BatchOleInsert
{
    public static void Main()
    {
        // Path to the Excel file that will be embedded as an OLE object.
        const string excelFilePath = "Sample.xlsx";

        // List of Word document file names to create and populate.
        List<string> wordFileNames = new List<string>
        {
            "Document1.docx",
            "Document2.docx",
            "Document3.docx"
        };

        // Ensure the Excel file exists; otherwise the example cannot run.
        if (!File.Exists(excelFilePath))
        {
            // Create a minimal Excel file for demonstration purposes.
            // In a real scenario the file would already exist.
            File.WriteAllBytes(excelFilePath, new byte[] { 0x50, 0x4B, 0x03, 0x04 }); // placeholder ZIP header
        }

        // Process each Word file.
        foreach (string wordFilePath in wordFileNames)
        {
            // Create a new blank Word document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph indicating which file is being processed.
            builder.Writeln($"This document contains an embedded Excel OLE object: {Path.GetFileName(excelFilePath)}");

            // Insert the Excel file as an embedded OLE object (not as an icon).
            // Using the overload that takes a file name: (fileName, isLinked, asIcon, presentation)
            builder.InsertOleObject(excelFilePath, false, false, null);

            // Save the document to the specified path.
            doc.Save(wordFilePath);
        }
    }
}
