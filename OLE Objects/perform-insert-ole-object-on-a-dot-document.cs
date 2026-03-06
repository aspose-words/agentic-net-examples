using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleIntoDot
{
    static void Main()
    {
        // Load an existing DOT template.
        Document doc = new Document("Template.dot");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a description before the OLE object.
        builder.Writeln("Embedded OLE package:");

        // Open the file that will be embedded as an OLE object.
        using (FileStream fileStream = File.Open("Data.zip", FileMode.Open, FileAccess.Read))
        {
            // Insert the OLE object from the stream.
            // progId "Package" indicates a generic OLE package.
            // asIcon = false inserts the object as its content.
            // presentation = null lets Aspose.Words choose a default icon if needed.
            Shape oleShape = builder.InsertOleObject(fileStream, "Package", false, null);

            // Optionally set the file name and display name for the OLE package.
            oleShape.OleFormat.OlePackage.FileName = "Data.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Data.zip";
        }

        // Save the modified document. The output can be DOCX, DOC, or another DOT.
        doc.Save("Result.docx");
    }
}
