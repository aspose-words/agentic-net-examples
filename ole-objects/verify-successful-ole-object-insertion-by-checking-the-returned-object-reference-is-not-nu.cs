using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary file that will be embedded as an OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "Sample.txt");
        File.WriteAllText(tempFilePath, "Sample content for OLE object.");

        // Initialize a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object (embedded, not as an icon) from the temporary file.
        // The method returns a Shape that contains the OLE object.
        Shape oleShape = builder.InsertOleObject(tempFilePath, "Package", false, false, null);

        // Verify that the insertion succeeded by checking the returned references.
        if (oleShape == null)
            throw new InvalidOperationException("InsertOleObject returned a null Shape.");

        OleFormat oleFormat = oleShape.OleFormat;
        if (oleFormat == null)
            throw new InvalidOperationException("OleFormat property is null.");

        // Optionally, you could inspect properties such as IsLink or OleIcon here.

        // Save the document to a temporary location to complete the lifecycle.
        string outputPath = Path.Combine(Path.GetTempPath(), "OleInsertionDemo.docx");
        doc.Save(outputPath);
    }
}
