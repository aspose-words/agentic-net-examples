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

        // Prepare a sample file to embed as an OLE object.
        string tempDir = Path.GetTempPath();
        string sampleFilePath = Path.Combine(tempDir, "SampleTextFile.txt");
        File.WriteAllText(sampleFilePath, "This is a sample text file embedded as an OLE icon.");

        // Insert the OLE object as an icon.
        // Passing null for iconFile and iconCaption makes Aspose.Words use the default system icon and file name as caption.
        builder.InsertOleObjectAsIcon(sampleFilePath, false, null, null);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OleObjectIcon.docx");
        doc.Save(outputPath);
    }
}
