using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary document and the extracted OLE binary file.
        const string docPath = "OleDocument.docx";
        const string extractedPath = "ExtractedOle.bin";

        // -------------------------------------------------
        // 1. Create a new document and embed a simple OLE package.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample data that will be stored inside the OLE object.
        byte[] oleData = System.Text.Encoding.UTF8.GetBytes("This is the content of the embedded OLE object.");

        using (MemoryStream oleStream = new MemoryStream(oleData))
        {
            // Insert the OLE object as a generic package.
            // Parameters: data stream, progId ("Package"), display as content (false), no custom icon (null).
            Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

            // Optional: give the package a file name and display name.
            oleShape.OleFormat.OlePackage.FileName = "Sample.txt";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample.txt";
        }

        // Save the document that now contains the OLE object.
        doc.Save(docPath);

        // -------------------------------------------------
        // 2. Load the document (demonstrating the load rule).
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Locate the first shape that holds an OLE object.
        Shape shape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        OleFormat oleFormat = shape.OleFormat;

        // -------------------------------------------------
        // 3. Extract the OLE object data to a binary file.
        // -------------------------------------------------
        using (FileStream fileStream = new FileStream(extractedPath, FileMode.Create))
        {
            // The OleFormat.Save method writes the embedded object's raw stream.
            oleFormat.Save(fileStream);
        }

        // Indicate that the operation has completed.
        Console.WriteLine("OLE object extracted to: " + Path.GetFullPath(extractedPath));
    }
}
