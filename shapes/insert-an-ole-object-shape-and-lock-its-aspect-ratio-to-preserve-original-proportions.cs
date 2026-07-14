using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class InsertOleObjectExample
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "OleObjectWithLockedAspectRatio.docx");

        // Create a simple text file to act as the OLE object source.
        string oleSourcePath = Path.Combine(artifactsDir, "Sample.txt");
        File.WriteAllText(oleSourcePath, "This is a sample OLE object content.");

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object. Use the text file as the source, embed it (isLinked = false),
        // display it as content (asIcon = false), and let Aspose.Words use the default presentation.
        Shape oleShape = builder.InsertOleObject(oleSourcePath, false, false, null);

        // Lock the aspect ratio to preserve original proportions.
        oleShape.AspectRatioLocked = true;

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        Console.WriteLine("Document created successfully at: " + outputPath);
    }
}
