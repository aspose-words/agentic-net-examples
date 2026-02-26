using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleIntoDot
{
    static void Main()
    {
        // Create a new empty document (DOT template)
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object
        string oleFilePath = @"C:\Data\SamplePackage.zip";

        // Path to a custom icon (ICO) that will represent the OLE object in the document
        string iconPath = @"C:\Images\PackageIcon.ico";

        // Insert the OLE object as an icon. The object is embedded (isLinked = false).
        // The icon caption will be the file name by default.
        builder.InsertOleObjectAsIcon(oleFilePath, false, iconPath, null);

        // Optionally add a paragraph after the OLE object
        builder.InsertParagraph();

        // Save the document as a DOT template
        string outputPath = @"C:\Output\TemplateWithOle.dot";
        doc.Save(outputPath, SaveFormat.Dot);
    }
}
