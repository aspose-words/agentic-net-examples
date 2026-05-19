using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Folder containing the source Word documents.
        string sourceFolder = "InputDocs";

        // Folder where the modified documents will be saved.
        string outputFolder = "OutputDocs";
        Directory.CreateDirectory(outputFolder);

        // Path to the Excel file that will be inserted as an OLE object.
        string excelFilePath = "Sample.xlsx";

        // Process each .docx file in the source folder.
        foreach (string docPath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the existing Word document.
            Document doc = new Document(docPath);

            // Create a DocumentBuilder to modify the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph describing the OLE object.
            builder.Writeln("Inserted Excel OLE object:");

            // Insert the Excel file as an embedded OLE object (not linked, not as an icon).
            // The overload InsertOleObject(string fileName, bool isLinked, bool asIcon, Stream presentation) is used.
            builder.InsertOleObject(excelFilePath, false, false, null);

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(docPath));
            doc.Save(outputPath);
        }
    }
}
