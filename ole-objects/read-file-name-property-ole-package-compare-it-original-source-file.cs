using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class OlePackageFileNameComparer
{
    static void Main()
    {
        // Path to the Word document that contains the OLE package.
        string documentPath = @"C:\Docs\DocumentWithOlePackage.docx";

        // Path to the original source file that was embedded as the OLE package.
        // This is the file we expect the OLE package to reference.
        string originalSourcePath = @"C:\SourceFiles\Archive.zip";

        // Load the document from disk.
        Document doc = new Document(documentPath);

        // Retrieve the original file name of the loaded document (if any).
        // This property is null when the document is loaded from a stream,
        // but we load from a file, so it will contain the full path.
        string docOriginalFileName = doc.OriginalFileName;

        Console.WriteLine($"Document original file name: {docOriginalFileName}");

        // Find the first shape that contains an OLE object.
        Shape oleShape = doc.GetChildNodes(NodeType.Shape, true)
                            .OfType<Shape>()
                            .FirstOrDefault(s => s.ShapeType == ShapeType.OleObject);

        if (oleShape == null)
        {
            Console.WriteLine("No OLE object found in the document.");
            return;
        }

        // Access the OleFormat of the shape.
        OleFormat oleFormat = oleShape.OleFormat;

        // Ensure the OLE object is an OLE Package.
        if (oleFormat?.OlePackage == null)
        {
            Console.WriteLine("The OLE object is not an OLE Package.");
            return;
        }

        // Read the file name stored inside the OLE package.
        string olePackageFileName = oleFormat.OlePackage.FileName;
        Console.WriteLine($"OLE Package file name property: {olePackageFileName}");

        // Compare the OLE package file name with the original source file name.
        // Use Path.GetFileName to compare only the file name part, ignoring directories.
        string expectedFileName = Path.GetFileName(originalSourcePath);
        bool namesMatch = string.Equals(olePackageFileName, expectedFileName, StringComparison.OrdinalIgnoreCase);

        Console.WriteLine($"Expected file name: {expectedFileName}");
        Console.WriteLine($"Do the names match? {(namesMatch ? "Yes" : "No")}");

        // Optionally, save the document after any modifications (none in this example).
        string outputPath = @"C:\Docs\DocumentWithOlePackage_Processed.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
