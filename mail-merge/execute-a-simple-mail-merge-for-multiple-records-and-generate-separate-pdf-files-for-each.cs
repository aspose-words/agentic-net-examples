using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

public class MailMergeMultiplePdf
{
    public static void Main()
    {
        // Create a template document with merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");
        builder.Writeln();

        // Prepare a DataTable with several records.
        DataTable data = new DataTable("Recipients");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");
        data.Rows.Add("John", "Doe", "Hello John!");
        data.Rows.Add("Jane", "Smith", "Welcome Jane!");
        data.Rows.Add("Bob", "Brown", "Good day Bob!");

        // Generate a separate PDF for each record.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the template so each PDF starts from the same base.
            Document doc = (Document)template.Clone(true);
            // Merge the current row into the document.
            doc.MailMerge.Execute(data.Rows[i]);
            // Save as PDF. The file name includes the record index.
            string pdfPath = $"MergedDocument_{i + 1}.pdf";
            doc.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"Saved: {pdfPath}");
        }
    }
}
