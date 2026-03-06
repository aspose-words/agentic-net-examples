using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleObjectIntoDotm
{
    static void Main()
    {
        // Path to the folder that contains the template and the OLE source file.
        string dataDir = @"C:\Data\";

        // Load an existing DOTM (macro‑enabled template) document.
        string templatePath = Path.Combine(dataDir, "Template.dotm");
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Open a stream that contains the data to be embedded as an OLE object.
        // In this example we embed an Excel spreadsheet.
        using (FileStream oleStream = File.Open(Path.Combine(dataDir, "Spreadsheet.xlsx"), FileMode.Open))
        {
            // Insert the OLE object at the current cursor position.
            // Parameters:
            //   oleStream   – stream with the application data.
            //   progId      – programmatic identifier of the OLE object ("Excel.Sheet.12").
            //   asIcon      – false to display the object content, true to display it as an icon.
            //   presentation– null to use the default presentation image.
            builder.InsertOleObject(oleStream, "Excel.Sheet.12", false, null);
        }

        // Save the modified document back as a DOTM file.
        string outputPath = Path.Combine(dataDir, "Result.dotm");
        doc.Save(outputPath);
    }
}
