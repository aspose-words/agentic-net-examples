using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class EnumerateOlePackages
{
    static void Main()
    {
        // Path to the Word document that may contain legacy OLE packages.
        // Update this path to point to an existing .docx file if you want to test with a real document.
        string docPath = @"C:\Docs\SampleWithOlePackage.docx";

        if (!File.Exists(docPath))
        {
            Console.WriteLine($"Document not found at '{docPath}'.");
            Console.WriteLine("Creating a temporary empty document for demonstration purposes.");

            // Create an empty document so the program can continue without throwing.
            Document emptyDoc = new Document();
            docPath = Path.GetTempFileName() + ".docx";
            emptyDoc.Save(docPath);
        }

        Document doc = new Document(docPath);

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Ensure the shape actually contains an OLE object.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Access the OlePackage; it will be null if the OLE object is not a Package.
            OlePackage olePackage = oleFormat.OlePackage;
            if (olePackage == null)
                continue;

            // Output the package's file name and display name.
            Console.WriteLine("Found OLE Package:");
            Console.WriteLine($"  FileName   : {olePackage.FileName}");
            Console.WriteLine($"  DisplayName: {olePackage.DisplayName}");

            // Optionally, you can extract the raw OLE data for further inspection.
            // byte[] rawData = oleFormat.GetRawData();
            // Console.WriteLine($"  Raw data length: {rawData.Length} bytes");
        }

        // Clean up the temporary file if we created one.
        if (docPath.EndsWith(".docx") && Path.GetTempPath().Equals(Path.GetDirectoryName(docPath) + Path.DirectorySeparatorChar.ToString(), StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                File.Delete(docPath);
            }
            catch
            {
                // Ignore any errors during cleanup.
            }
        }
    }
}
