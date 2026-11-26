using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a temporary text file that will be embedded as an OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleOle.txt");
        File.WriteAllText(tempFilePath, "This is sample OLE content.");

        // Insert the OLE object (embedded, not as an icon) at the current cursor position.
        // Parameters: file name, isLinked = false (embed), asIcon = false (show content), presentation = null.
        Shape oleShape = builder.InsertOleObject(tempFilePath, false, false, null);

        // Lock the aspect ratio to preserve the original proportions of the OLE object.
        oleShape.AspectRatioLocked = true;

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectShape.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Clean up the temporary file.
        File.Delete(tempFilePath);
    }
}
