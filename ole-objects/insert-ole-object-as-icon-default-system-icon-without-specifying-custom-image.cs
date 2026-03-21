using Aspose.Words;
using System;
using System.IO;

class InsertOleIconExample
{
    static void Main()
    {
        // Create a temporary file to embed as an OLE object.
        string tempDir = Path.GetTempPath();
        string oleFilePath = Path.Combine(tempDir, "Sample.txt");
        File.WriteAllText(oleFilePath, "This is a sample text file used for OLE embedding.");

        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object as an icon.
        // - isLinked = false (embed the file).
        // - iconFile = null (use the default system icon for the file type).
        // - iconCaption = null (use the file name as the caption).
        builder.InsertOleObjectAsIcon(oleFilePath, false, null, null);

        // Save the resulting document to the temporary folder.
        string outputPath = Path.Combine(tempDir, "OleIcon.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
