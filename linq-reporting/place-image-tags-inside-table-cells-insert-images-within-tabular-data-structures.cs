using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // DocumentBuilder provides methods to construct the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        builder.StartTable();

        // ---------- First Row ----------
        // Cell (0,0) – insert an image.
        builder.InsertCell();
        // InsertImage returns a Shape object that can be further configured.
        Shape image1 = builder.InsertImage(@"C:\Images\Image1.jpg");
        // Set explicit size for the image (optional).
        image1.Width = 100;   // points
        image1.Height = 100;  // points

        // Cell (0,1) – insert some text.
        builder.InsertCell();
        builder.Write("First image description");

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Cell (1,0) – insert text.
        builder.InsertCell();
        builder.Write("Second image description");

        // Cell (1,1) – insert another image.
        builder.InsertCell();
        Shape image2 = builder.InsertImage(@"C:\Images\Image2.png");
        image2.Width = 80;
        image2.Height = 80;

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save(@"C:\Output\TableWithImages.docx");
    }
}
