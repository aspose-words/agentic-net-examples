using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary text file to embed as an OLE object.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeOleExample");
        Directory.CreateDirectory(tempFolder);
        string oleFilePath = Path.Combine(tempFolder, "Sample.txt");
        File.WriteAllText(oleFilePath, "This is a sample text file embedded as an OLE object.");

        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object as an icon.
        // Parameters: file name, isLinked (false = embed), iconFile (null = default system icon), iconCaption (null = use file name).
        builder.InsertOleObjectAsIcon(oleFilePath, false, null, null);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectAsIcon.docx");
        doc.Save(outputPath);
    }
}
