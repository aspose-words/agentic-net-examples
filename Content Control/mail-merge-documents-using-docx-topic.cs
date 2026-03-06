using System;
using System.Data;
using Aspose.Words;

class MailMergeExample
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a simple DOCX template with merge fields.
        // -------------------------------------------------
        Document template = new Document();                     // create a new document
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.InsertField(" MERGEFIELD CustomerName ");       // first merge field
        builder.InsertParagraph();                              // line break
        builder.InsertField(" MERGEFIELD Address ");            // second merge field

        // Save the template to disk (optional, shows the create‑save lifecycle).
        template.Save("Template.docx");                         // save the document

        // -------------------------------------------------
        // 2. Load the template for mail merge.
        // -------------------------------------------------
        Document doc = new Document("Template.docx");           // load the existing DOCX

        // -------------------------------------------------
        // 3. Prepare a data source (DataTable) for the merge.
        // -------------------------------------------------
        DataTable table = new DataTable("Customers");
        table.Columns.Add("CustomerName");
        table.Columns.Add("Address");
        table.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
        table.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");

        // -------------------------------------------------
        // 4. Execute the mail merge using the DataTable.
        // -------------------------------------------------
        doc.MailMerge.Execute(table);                          // perform merge for all rows

        // -------------------------------------------------
        // 5. Save the merged result.
        // -------------------------------------------------
        doc.Save("Merged.docx");                               // save the final document
    }
}
