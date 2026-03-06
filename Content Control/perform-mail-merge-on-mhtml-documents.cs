using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class MailMergeMhtmlExample
{
    static void Main()
    {
        // Path to the source MHTML document that contains MERGEFIELD tags.
        string mhtmlPath = @"C:\Docs\Template.mht";

        // Load the MHTML document.
        Document doc = new Document(mhtmlPath);

        // Prepare a data source for the mail merge.
        // The column names must match the MERGEFIELD names in the template.
        DataTable data = new DataTable("Customer");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Address");
        data.Rows.Add("John", "Doe", "123 Main St, Anytown");
        data.Rows.Add("Jane", "Smith", "456 Oak Ave, Othertown");

        // Execute the mail merge. This will replace the MERGEFIELDs with the data from the DataTable.
        // Since the template does not contain mail‑merge regions, we use the simple Execute method.
        doc.MailMerge.Execute(data);

        // Save the merged document. It can be saved back to MHTML or any other supported format.
        string outputPath = @"C:\Docs\MergedResult.docx";
        doc.Save(outputPath);
    }
}
