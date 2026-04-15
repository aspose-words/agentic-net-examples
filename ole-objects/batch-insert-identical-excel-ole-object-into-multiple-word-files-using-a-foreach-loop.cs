using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class BatchOleInsert
{
    public static void Main()
    {
        // Folder where the Excel file to embed is located.
        string excelFilePath = Path.Combine(Environment.CurrentDirectory, "SampleData.xlsx");

        // Ensure the Excel file exists; create a dummy file if it does not.
        if (!File.Exists(excelFilePath))
        {
            // Create a minimal Excel file (empty ZIP with .xlsx extension) for demonstration.
            File.WriteAllBytes(excelFilePath, new byte[] { 80, 75, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 });
        }

        // List of output Word document file names.
        List<string> outputDocs = new List<string>
        {
            "DocumentA.docx",
            "DocumentB.docx",
            "DocumentC.docx"
        };

        // Process each document: create, insert the same Excel OLE object, and save.
        foreach (string fileName in outputDocs)
        {
            // Create a new blank Word document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a paragraph indicating the inserted OLE object.
            builder.Writeln($"This document contains an embedded Excel OLE object ({Path.GetFileName(excelFilePath)}):");

            // Insert the Excel file as an embedded OLE object (not as an icon).
            // Parameters: file name, isLinked = false, asIcon = false, presentation = null.
            builder.InsertOleObject(excelFilePath, false, false, null);

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, fileName);
            doc.Save(outputPath);
        }
    }
}
