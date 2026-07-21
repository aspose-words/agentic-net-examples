using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class InsertPictureInTableCell
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert the first cell.
        builder.StartTable();
        builder.InsertCell();

        // A tiny red PNG image (1x1 pixel) encoded in Base64.
        // This avoids the need for System.Drawing or external image files.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lK5XAAAAAElFTkSuQmCC";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Insert the image into the cell as a floating shape.
        // Use the overload that accepts a byte array and then set the desired size.
        Shape pictureShape = builder.InsertImage(imageBytes);
        pictureShape.Width = 100;   // width in points
        pictureShape.Height = 100;  // height in points

        // Enable layout inside the cell so the shape moves with the cell.
        pictureShape.IsLayoutInCell = true;
        pictureShape.WrapType = WrapType.None;

        // Finish the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PictureInTableCell.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");
    }
}
