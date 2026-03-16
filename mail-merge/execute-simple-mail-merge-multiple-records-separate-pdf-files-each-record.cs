using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class MailMergeSeparatePdfs
{
    static void Main()
    {
        // Create a mail‑merge template in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.InsertField(" MERGEFIELD FullName ");
        builder.InsertParagraph();
        builder.InsertField(" MERGEFIELD Address ");

        // Prepare a DataTable that holds several records.
        DataTable data = new DataTable("Customers");
        data.Columns.Add("FullName");
        data.Columns.Add("Address");
        data.Rows.Add(new object[] { "Thomas Hardy", "120 Hanover Sq., London" });
        data.Rows.Add(new object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });
        data.Rows.Add(new object[] { "John Doe", "123 Main St., New York" });

        // Iterate over each DataRow, perform a single‑record mail merge,
        // and save the result as an individual PDF file.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the template so that each iteration works with a fresh copy.
            Document doc = (Document)template.Clone();

            // Execute mail merge for the current row only.
            doc.MailMerge.Execute(data.Rows[i]);

            // Save the merged document as PDF. File name includes the record index.
            string outputPath = $"Customer_{i + 1}.pdf";
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
