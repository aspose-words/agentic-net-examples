using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the file that will be embedded as an OLE object (e.g., an Excel workbook).
        string sourceFilePath = @"C:\Data\Sample.xlsx";

        // Open the source file as a stream.
        using (FileStream sourceStream = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object.
            // progId "Excel.Sheet" identifies the OLE type.
            // asIcon = false inserts the object as its content (not as an icon).
            // presentation = null lets Aspose.Words use a default preview image.
            Shape oleShape = builder.InsertOleObject(sourceStream, "Excel.Sheet", false, null);
        }

        // Save the resulting document in DOCX format.
        string outputPath = @"C:\Output\InsertOleObject.docx";
        doc.Save(outputPath);
    }
}
