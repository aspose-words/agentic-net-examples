using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToPngExample
{
    static void Main()
    {
        // Load the source DOCX document.
        // The Document constructor automatically detects the format from the file extension.
        Document doc = new Document("Template.docx");

        // -------------------------------------------------
        // 1. Mail merge using an array of field names and values (single record).
        // -------------------------------------------------
        string[] fieldNames = { "FirstName", "LastName", "Address" };
        object[] fieldValues = { "John", "Doe", "120 Hanover Sq., London" };
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // -------------------------------------------------
        // 2. Mail merge using a DataTable (multiple records).
        // -------------------------------------------------
        DataTable table = new DataTable("Customers");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Columns.Add("Address");

        table.Rows.Add("Alice", "Smith", "10 Downing St., London");
        table.Rows.Add("Bob", "Johnson", "1600 Pennsylvania Ave., Washington");

        // The Execute method will repeat the whole document for each row in the table.
        doc.MailMerge.Execute(table);

        // -------------------------------------------------
        // Save the merged document as PNG.
        // Each page of the document will be rendered to a separate PNG file.
        // -------------------------------------------------
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render all pages; the PageSet property can be used to limit pages.
            PageSet = new PageSet(0) // 0 means start from the first page.
        };

        // Save the first page as "Result.png". If the document has multiple pages,
        // you can loop over doc.PageCount and change the PageSet accordingly.
        doc.Save("Result.png", options);
    }
}
