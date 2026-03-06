using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleIntoDot
{
    static void Main()
    {
        // Paths to the template, the file to embed, and the output document.
        string dataDir = @"C:\Data\";
        string templatePath = Path.Combine(dataDir, "Template.dot");
        string oleFilePath = Path.Combine(dataDir, "cat001.zip"); // any file to embed
        string outputPath = Path.Combine(dataDir, "Result.dot");

        // Load the existing DOT template.
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an embedded OLE object (Package) from a stream.
        // Parameters: stream, progId ("Package" for generic files), asIcon = false, presentation = null.
        using (FileStream oleStream = new FileStream(oleFilePath, FileMode.Open, FileAccess.Read))
        {
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // Optional: set the file name and display name that Word will show.
            oleShape.OleFormat.OlePackage.FileName = Path.GetFileName(oleFilePath);
            oleShape.OleFormat.OlePackage.DisplayName = "Sample ZIP";
        }

        // Save the modified document back as a DOT file.
        doc.Save(outputPath, SaveFormat.Dot);
    }
}
