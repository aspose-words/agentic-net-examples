using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectIntoDotx
{
    static void Main()
    {
        // Path to the DOTX template file.
        string templatePath = @"C:\Docs\Template.dotx";

        // Load the DOTX document.
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph to separate the OLE object from existing content.
        builder.Writeln("Embedded Excel Spreadsheet:");

        // Path to the file that will be embedded as an OLE object.
        string oleFilePath = @"C:\Docs\SampleData.xlsx";

        // Insert the OLE object.
        // Parameters:
        //   fileName   – full path to the file to embed.
        //   isLinked   – false to embed (not link) the object.
        //   asIcon     – false to display the object content (set true to show as an icon).
        //   presentation – null to use the default presentation image.
        builder.InsertOleObject(oleFilePath, false, false, null);

        // Save the resulting document as a regular DOCX file.
        string outputPath = @"C:\Docs\Result.docx";
        doc.Save(outputPath);
    }
}
