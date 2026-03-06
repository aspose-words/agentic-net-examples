using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Load the MHTML template that contains a mail‑merge region.
        // The template must have MERGEFIELD tags like:
        //   MERGEFIELD TableStart:Products
        //   MERGEFIELD Name
        //   MERGEFIELD Price
        //   MERGEFIELD TableEnd:Products
        Document doc = new Document("Template.mhtml");

        // Prepare a DataTable that matches the region name ("Products").
        DataTable products = new DataTable("Products");
        products.Columns.Add("Name");
        products.Columns.Add("Price");

        // Add sample rows – the mail merge will repeat the table row for each record.
        products.Rows.Add("Apple", "$1.20");
        products.Rows.Add("Banana", "$0.80");
        products.Rows.Add("Cherry", "$2.50");

        // Execute the mail merge using regions. This will replace the MERGEFIELDs inside the
        // TableStart/TableEnd block with the data from the DataTable, creating a row for each record.
        doc.MailMerge.ExecuteWithRegions(products);

        // Save the result. The output can be any format supported by Aspose.Words.
        doc.Save("Result.docx");
    }
}
