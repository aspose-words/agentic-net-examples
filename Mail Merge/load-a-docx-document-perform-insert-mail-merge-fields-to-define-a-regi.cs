using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        // The constructor automatically detects the format.
        Document doc = new Document("Input.docx");

        // Use DocumentBuilder to insert a mail‑merge region.
        // A region is defined by a TableStart field and a matching TableEnd field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentStart();                     // Position at the beginning of the document.

        // Insert the start of the region named "MyRegion".
        builder.InsertField(" MERGEFIELD TableStart:MyRegion ");

        // Insert a sample merge field that will be repeated inside the region.
        builder.Write("Name: ");
        builder.InsertField(" MERGEFIELD Name ");

        // Insert the end of the region.
        builder.InsertField(" MERGEFIELD TableEnd:MyRegion ");

        // (Optional) Perform a mail merge with a dummy data source so the document has content.
        // Here we create a simple DataTable with one column "Name".
        DataTable table = new DataTable("MyRegion");
        table.Columns.Add("Name");
        table.Rows.Add("Alice");
        table.Rows.Add("Bob");
        doc.MailMerge.ExecuteWithRegions(table);

        // Save the resulting document as PNG images.
        // Each page will be saved as a separate PNG file (e.g., Output.png, Output_1.png, ...).
        doc.Save("Output.png", SaveFormat.Png);
    }
}
