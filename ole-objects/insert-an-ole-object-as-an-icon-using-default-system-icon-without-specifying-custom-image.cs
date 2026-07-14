using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a temporary file to embed as an OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleText.txt");
        File.WriteAllText(tempFilePath, "This is a sample text file embedded as an OLE object.");

        // Insert the OLE object as an icon.
        // Parameters: file name, isLinked = false (embed), iconFile = null (use default system icon), iconCaption = null (use file name).
        builder.InsertOleObjectAsIcon(tempFilePath, false, null, null);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectAsIcon.docx");
        doc.Save(outputPath);
    }
}
