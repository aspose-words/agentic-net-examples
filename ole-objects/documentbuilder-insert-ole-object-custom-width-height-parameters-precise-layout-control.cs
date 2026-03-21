using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertOleWithSize
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a temporary file to embed as an OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleData.txt");
        File.WriteAllText(tempFilePath, "This is sample embedded content.");

        // Insert the OLE object. This overload returns a Shape that contains the OLE data.
        // Parameters:
        //   fileName – full path to the file.
        //   progId   – ProgID for the OLE object; "Package" works for generic files.
        //   isLinked – false = embed the object, true = link to the file.
        //   asIcon   – false = display the object's content, true = display as an icon.
        //   presentation – null = use default icon.
        Shape oleShape = builder.InsertOleObject(
            tempFilePath,
            "Package",
            false,   // embed, not link
            false,   // display content, not icon
            null);   // default presentation

        // Set custom size for precise layout control.
        // Width and Height are measured in points (1 point = 1/72 inch).
        oleShape.Width = 300;   // 300 points ≈ 4.17 inches
        oleShape.Height = 200;  // 200 points ≈ 2.78 inches

        // Position the OLE object relative to the page margins.
        oleShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        oleShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        oleShape.Left = 100;
        oleShape.Top = 100;
        oleShape.WrapType = WrapType.Square; // Text will wrap around the object.

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithOle.docx");
        doc.Save(outputPath);

        // Clean up the temporary file.
        File.Delete(tempFilePath);
    }
}
