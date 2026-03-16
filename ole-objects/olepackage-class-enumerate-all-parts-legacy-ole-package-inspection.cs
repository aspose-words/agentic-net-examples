using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class EnumerateOlePackages
{
    static void Main()
    {
        // Load the Word document that may contain legacy OLE packages.
        // Replace the path with the actual location of your document.
        string docPath = @"C:\Docs\SampleWithOlePackage.docx";
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
    }
}
