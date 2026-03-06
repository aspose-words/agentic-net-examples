using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Directory that contains the source file and optional icon.
        string dataDir = @"C:\Data\";

        // Path to the file that will be embedded as an OLE object.
        string oleFilePath = Path.Combine(dataDir, "Spreadsheet.xlsx");

        // Insert the OLE object as embedded content (displayed as the actual data, not as an icon).
        // Overload: InsertOleObject(string fileName, bool isLinked, bool asIcon, Stream presentation)
        // Parameters: fileName, isLinked = false (embed), asIcon = false (show content), presentation = null (use default icon if needed).
        builder.InsertOleObject(oleFilePath, false, false, null);

        // Insert the same OLE object as an icon with a custom caption and custom icon image.
        // Overload: InsertOleObjectAsIcon(string fileName, bool isLinked, string iconFile, string iconCaption)
        string iconPath = Path.Combine(dataDir, "Icon.ico");
        builder.InsertOleObjectAsIcon(oleFilePath, false, iconPath, "My Excel Sheet");

        // Save the resulting document.
        string outputPath = @"C:\Output\InsertOleObject.docx";
        doc.Save(outputPath);
    }
}
