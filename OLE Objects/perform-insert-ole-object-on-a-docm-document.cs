using System.IO;
using Aspose.Words;

class InsertOleObjectExample
{
    static void Main()
    {
        // Path to the folder that contains the source DOCM and the file to embed.
        string dataDir = @"C:\Data\";

        // Load an existing DOCM document.
        Document doc = new Document(Path.Combine(dataDir, "Template.docm"));

        // Create a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph before the OLE object (optional).
        builder.Writeln("Embedded Excel spreadsheet:");

        // Open the file that will be embedded as an OLE object.
        using (FileStream excelStream = new FileStream(Path.Combine(dataDir, "Sample.xlsx"), FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object.
            // Parameters:
            //   stream      – the data stream of the file to embed.
            //   progId      – programmatic identifier of the OLE object (Excel sheet in this case).
            //   asIcon      – false to display the content, true to display as an icon.
            //   presentation– null to use the default presentation image.
            builder.InsertOleObject(excelStream, "Excel.Sheet.12", false, null);
        }

        // Save the modified document as a DOCM file.
        doc.Save(Path.Combine(dataDir, "Result.docm"));
    }
}
