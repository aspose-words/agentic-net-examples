using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Folder that contains the Word documents to be processed
        string docsFolder = @"C:\Docs";

        // Path to the Excel file that will be embedded as an OLE object
        string excelFile = @"C:\Data\Sample.xlsx";

        // Retrieve all .docx files in the specified folder
        string[] wordFiles = Directory.GetFiles(docsFolder, "*.docx");

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
    }
}
