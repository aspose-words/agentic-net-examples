using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // 1. Create a new blank document (lifecycle rule: create)
        Document doc = new Document();

        // 2. Attach a DocumentBuilder to the document (required for inserting content)
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 3. Format a numbered list
        // -------------------------------------------------
        // Create a default numbered list template
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        // Apply the list to subsequent paragraphs
        builder.ListFormat.List = list;
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        // Stop list formatting
        builder.ListFormat.RemoveNumbers();

        builder.Writeln(); // blank line between sections

        // -------------------------------------------------
        // 4. Insert an OfficeMath equation using a field
        // -------------------------------------------------
        // The EQ field renders a simple equation: a² + b² = c²
        builder.InsertField(@"EQ \\o\\ac( a,2 ) + \\o\\ac( b,2 ) = \\o\\ac( c,2 )");

        builder.Writeln(); // blank line

        // -------------------------------------------------
        // 5. Insert a SmartArt shape and trigger cold rendering
        // -------------------------------------------------
        // Insert a rectangle as a placeholder for SmartArt
        Shape smartArtShape = builder.InsertShape(ShapeType.Rectangle, 200, 100);
        // Force the SmartArt rendering engine to update (cold rendering)
        smartArtShape.UpdateSmartArtDrawing();

        builder.Writeln(); // blank line

        // -------------------------------------------------
        // 6. Generate a custom barcode image and insert it
        // -------------------------------------------------
        byte[] barcodeBytes = GetPlaceholderBarcodeBytes();
        // Insert the barcode image inline, using the same dimensions as generated
        builder.InsertImage(barcodeBytes, 300, 100);

        // -------------------------------------------------
        // 7. Save the document (lifecycle rule: save)
        // -------------------------------------------------
        doc.Save("FormattedList_OfficeMath_SmartArt_Barcode.docx");
    }

    // Helper method: returns a simple placeholder PNG that represents a barcode.
    // In a real scenario you could replace this with actual barcode generation logic.
    static byte[] GetPlaceholderBarcodeBytes()
    {
        // This is a 1x1 black pixel PNG encoded in base64 – serves as a minimal placeholder.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        return Convert.FromBase64String(base64Png);
    }
}
