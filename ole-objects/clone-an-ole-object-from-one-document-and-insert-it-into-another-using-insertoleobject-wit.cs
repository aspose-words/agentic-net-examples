using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class CloneOleObjectExample
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create a simple file that will be embedded as an OLE object.
        string sampleFilePath = Path.Combine(workDir, "sample.txt");
        File.WriteAllText(sampleFilePath, "This is sample text for the OLE object.");

        // -----------------------------------------------------------------
        // 1. Create a source document and embed the sample file as an OLE object.
        // -----------------------------------------------------------------
        string sourceDocPath = Path.Combine(workDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        srcBuilder.Writeln("Source document with an embedded OLE object:");
        using (FileStream fs = File.OpenRead(sampleFilePath))
        {
            // Insert the OLE object as a package (generic container).
            srcBuilder.InsertOleObject(fs, "Package", false, null);
        }

        sourceDoc.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // 2. Load the source document and extract the OLE object's raw data.
        // -----------------------------------------------------------------
        Document loadedSource = new Document(sourceDocPath);
        // Find the first shape that contains an OLE object.
        Shape oleShape = (Shape)loadedSource.GetChild(NodeType.Shape, 0, true);
        OleFormat oleFormat = oleShape.OleFormat;

        // Save the OLE data into a memory stream.
        MemoryStream oleDataStream = new MemoryStream();
        oleFormat.Save(oleDataStream);
        oleDataStream.Position = 0; // Reset for reading.

        // Preserve the original ProgID (required for insertion).
        string progId = oleFormat.ProgId;

        // -----------------------------------------------------------------
        // 3. Create a target document and insert the cloned OLE object.
        // -----------------------------------------------------------------
        string targetDocPath = Path.Combine(workDir, "Target.docx");
        Document targetDoc = new Document();
        DocumentBuilder tgtBuilder = new DocumentBuilder(targetDoc);

        tgtBuilder.Writeln("Target document with the cloned OLE object:");
        // Insert the previously extracted OLE data.
        tgtBuilder.InsertOleObject(oleDataStream, progId, false, null);

        targetDoc.Save(targetDocPath);

        // Inform the user where the files are located.
        Console.WriteLine("Source document saved to: " + sourceDocPath);
        Console.WriteLine("Target document saved to: " + targetDocPath);
    }
}
