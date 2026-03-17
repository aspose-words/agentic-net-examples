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

        // Path to the file that will be embedded as an OLE object.
        // Replace with the actual file you want to embed.
        string oleFilePath = @"C:\Data\Sample.xlsx";

        // Insert the OLE object into the document.
        // Parameters:
        //   fileName   – full path to the file to embed.
        //   isLinked   – false to embed the object (true would create a link).
        //   asIcon     – false to display the object content (true would show an icon).
        //   presentation – null to use the default presentation image.
        Shape oleShape = builder.InsertOleObject(oleFilePath, false, false, null);

        // Lock the shape's aspect ratio so it preserves its original proportions.
        oleShape.AspectRatioLocked = true;

        // (Optional) Set explicit size for the shape in points.
        oleShape.Width = 200;   // width in points
        oleShape.Height = 150;  // height in points

        // Save the resulting document.
        doc.Save(@"C:\Output\OleObjectLocked.docx");
    }
}
