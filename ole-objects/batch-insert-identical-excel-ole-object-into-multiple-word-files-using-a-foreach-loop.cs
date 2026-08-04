using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path to the Excel file that will be embedded as an OLE object.
        string excelFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.xlsx");

        // Ensure the Excel file exists – create a minimal placeholder if it does not.
        if (!File.Exists(excelFilePath))
        {
            // A simple empty XLSX file (ZIP container with minimal structure) can be created.
            // For demonstration we just write a few bytes; Aspose.Words will still embed it.
            File.WriteAllBytes(excelFilePath, new byte[] { 0x50, 0x4B, 0x03, 0x04 });
        }

        // List of Word documents to process.
        List<string> wordFilePaths = new List<string>
        {
            Path.Combine(Directory.GetCurrentDirectory(), "Document1.docx"),
            Path.Combine(Directory.GetCurrentDirectory(), "Document2.docx"),
            Path.Combine(Directory.GetCurrentDirectory(), "Document3.docx")
        };

        foreach (string wordPath in wordFilePaths)
        {
            Document doc;

            // Load the document if it exists; otherwise create a new blank document and save it first.
            if (File.Exists(wordPath))
            {
                doc = new Document(wordPath);
            }
            else
            {
                doc = new Document();          // Create a blank document.
                doc.Save(wordPath);            // Persist the new document so the path is valid.
            }

            // Insert the Excel OLE object.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Embedded Excel OLE object:");
            // InsertOleObject(string fileName, bool isLinked, bool asIcon, Stream presentation)
            builder.InsertOleObject(excelFilePath, false, false, null);

            // Overwrite the original file with the modified document.
            doc.Save(wordPath);
        }
    }
}
