using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class OlePackageExample
{
    static void Main()
    {
        // Load an existing Word document that contains an OLE Package.
        // This uses the Document(string) constructor – a provided lifecycle rule.
        Document doc = new Document("InputWithOlePackage.docx");

        // Iterate through all shapes in the document to find the OLE Package.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Ensure the shape actually contains OLE data.
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue;

            // Access the OlePackage; it will be null if the OLE object is not a Package.
            OlePackage olePackage = oleFormat.OlePackage;
            if (olePackage == null)
                continue;

            // Read current properties.
            Console.WriteLine("Original FileName: " + olePackage.FileName);
            Console.WriteLine("Original DisplayName: " + olePackage.DisplayName);

            // Modify the properties.
            olePackage.FileName = "NewPackageFileName.zip";
            olePackage.DisplayName = "New Package Display Name.zip";

            Console.WriteLine("Modified FileName: " + olePackage.FileName);
            Console.WriteLine("Modified DisplayName: " + olePackage.DisplayName);
        }

        // Save the modified document.
        // This uses the Document.Save(string) method – a provided lifecycle rule.
        doc.Save("OutputWithModifiedOlePackage.docx");
    }
}
