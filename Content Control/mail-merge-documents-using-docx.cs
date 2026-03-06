using System;
using System.Data;
using Aspose.Words;

class MailMergeExample
{
    static void Main()
    {
        // ---------- Create a mail‑merge source document in memory ----------
        Document sourceDoc = new Document();                     // create
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.InsertField(" MERGEFIELD CustomerName ");        // first merge field
        builder.InsertParagraph();                               // line break
        builder.InsertField(" MERGEFIELD Address ");             // second merge field

        // ---------- Prepare data for the merge ----------
        DataTable data = new DataTable("Customers");
        data.Columns.Add("CustomerName");
        data.Columns.Add("Address");
        data.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
        data.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");

        // ---------- Execute the mail merge ----------
        sourceDoc.MailMerge.Execute(data);

        // ---------- Save the merged document ----------
        sourceDoc.Save("MergedOutput.docx");                     // save

        // ---------- Load an existing template, merge, and save ----------
        Document templateDoc = new Document("Template.docx");    // load
        // The template is expected to contain MERGEFIELD tags matching the column names.
        templateDoc.MailMerge.Execute(data);
        templateDoc.Save("MergedFromTemplate.docx");             // save
    }
}
