using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleIntoDotx
{
    static void Main()
    {
        // Paths to the template, OLE source file and optional icon image.
        string dataDir = @"C:\Data\";
        string templatePath = Path.Combine(dataDir, "Template.dotx");
        string oleFilePath = Path.Combine(dataDir, "Spreadsheet.xlsx");
        string iconPath = Path.Combine(dataDir, "Icon.png");
        string outputPath = Path.Combine(dataDir, "Result.docx");

        // Load the DOTX template.
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a caption before the OLE object.
        builder.Writeln("Embedded Excel spreadsheet:");

        // Insert the OLE object.
        // Parameters:
        //   oleFilePath – path to the file to embed.
        //   isLinked    – false for an embedded object.
        //   asIcon      – false to display the content, true to display as an icon.
        //   presentation– stream with a custom icon image (optional).
        using (FileStream iconStream = File.OpenRead(iconPath))
        {
            // Here we embed the Excel file and display it as an icon using a custom image.
            builder.InsertOleObject(oleFilePath, false, true, iconStream);
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
