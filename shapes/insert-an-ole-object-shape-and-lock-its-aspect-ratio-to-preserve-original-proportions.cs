using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class InsertOleObjectExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a temporary file that will be embedded as an OLE object.
        // Here we use a simple text file, but any file type can be used.
        string tempOleFile = Path.Combine(Path.GetTempPath(), "SampleOleObject.txt");
        File.WriteAllText(tempOleFile, "This is a sample OLE object content.");

        // Insert the OLE object into the document.
        // Parameters: file name, isLinked = false (embed), asIcon = false (show content), presentation = null.
        Shape oleShape = builder.InsertOleObject(tempOleFile, false, false, null);

        // Lock the aspect ratio of the OLE object shape to preserve its original proportions.
        oleShape.AspectRatioLocked = true;

        // Save the document to the current directory.
        string outputPath = "OleObjectWithLockedAspectRatio.docx";
        doc.Save(outputPath);

        // Validation: ensure the output file exists.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        // Validation: ensure the inserted shape is an OLE object and its aspect ratio is locked.
        Shape insertedOleShape = doc.GetChildNodes(NodeType.Shape, true)
                                    .OfType<Shape>()
                                    .FirstOrDefault(s => s.ShapeType == ShapeType.OleObject);
        if (insertedOleShape == null)
            throw new InvalidOperationException("The OLE object shape was not found in the document.");

        if (!insertedOleShape.AspectRatioLocked)
            throw new InvalidOperationException("The AspectRatioLocked property was not set correctly.");

        // Clean up the temporary OLE file.
        File.Delete(tempOleFile);
    }
}
