using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OleInsertionVerification
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a temporary file to act as the OLE source.
        string tempOlePath = Path.Combine(Path.GetTempPath(), $"Sample_{Guid.NewGuid()}.xlsx");
        File.WriteAllText(tempOlePath, "Dummy content for OLE object.");

        try
        {
            // Insert the OLE object. The method returns a Shape that contains the OLE data.
            Shape oleShape = builder.InsertOleObject(tempOlePath, isLinked: false, asIcon: false, presentation: null);

            // Verify that the returned Shape reference is not null.
            if (oleShape == null)
                throw new InvalidOperationException("InsertOleObject returned a null Shape.");

            // Verify that the Shape actually contains an OLE object.
            if (oleShape.OleFormat == null)
                throw new InvalidOperationException("The inserted Shape does not contain an OleFormat.");

            // Output some properties to confirm successful insertion.
            Console.WriteLine($"OLE object inserted. Shape type: {oleShape.ShapeType}");
            Console.WriteLine($"Is linked: {oleShape.OleFormat.IsLink}");
            Console.WriteLine($"Displayed as icon: {oleShape.OleFormat.OleIcon}");

            // Save the document to verify that the OLE object persists.
            string outputPath = Path.Combine(Path.GetTempPath(), $"OleInsertionResult_{Guid.NewGuid()}.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
        finally
        {
            // Clean up the temporary OLE file.
            if (File.Exists(tempOlePath))
                File.Delete(tempOlePath);
        }
    }
}
