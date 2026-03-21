using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a temporary working directory
        string workDir = Path.Combine(Path.GetTempPath(), "OleDemo");
        Directory.CreateDirectory(workDir);

        // Path to the Excel file that will be embedded as an OLE object
        string excelFile = Path.Combine(workDir, "Sample.xlsx");

        // Ensure a dummy Excel file exists (a minimal ZIP header is enough for the demo)
        if (!File.Exists(excelFile))
        {
            // Write a minimal ZIP file header so the file is recognized as an Office Open XML package
            byte[] minimalZipHeader = new byte[] { 0x50, 0x4B, 0x03, 0x04 };
            File.WriteAllBytes(excelFile, minimalZipHeader);
        }

        // Folder that contains the Word documents to be processed
        string docsFolder = workDir;

        // Ensure at least one .docx file exists for processing
        string[] wordFiles = Directory.GetFiles(docsFolder, "*.docx");
        if (wordFiles.Length == 0)
        {
            string sampleDoc = Path.Combine(docsFolder, "Sample.docx");
            var emptyDoc = new Document();
            emptyDoc.Save(sampleDoc);
            wordFiles = new[] { sampleDoc };
        }

        foreach (string wordFile in wordFiles)
        {
            // Load the existing Word document
            Document doc = new Document(wordFile);

            // Create a DocumentBuilder to modify the document
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Position the builder at the end of the document (adjust as needed)
            builder.MoveToDocumentEnd();

            // Optional: add a paragraph break before inserting the OLE object
            builder.InsertParagraph();

            // Insert the Excel OLE object:
            //   - fileName: path to the Excel file
            //   - isLinked: false (embed the file, not a link)
            //   - asIcon: false (display the content, not an icon)
            //   - presentation: null (use default presentation)
            builder.InsertOleObject(excelFile, false, false, null);

            // Save the modified document, overwriting the original file
            doc.Save(wordFile);
        }

        Console.WriteLine("OLE objects inserted successfully.");
    }
}
