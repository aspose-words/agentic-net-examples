using System;
using System.Data;
using Aspose.Words;

class MailMergeMhtmlExample
{
    static void Main()
    {
        // Load the source MHTML document that contains MERGEFIELD fields.
        // The Document constructor handles loading; no custom loading code is required.
        Document doc = new Document("SourceDocument.mht");

        // Prepare a simple data source for the mail merge.
        DataTable data = new DataTable("CustomerData");
        data.Columns.Add("Name");
        data.Columns.Add("Address");
        data.Rows.Add("John Doe", "123 Main St, Anytown");
        data.Rows.Add("Jane Smith", "456 Oak Ave, Othertown");

        // Execute the mail merge. This will replace each MERGEFIELD in the document
        // with the corresponding value from the DataTable.
        doc.MailMerge.Execute(data);

        // Save the merged document. The Save method determines the format from the file extension.
        doc.Save("MergedResult.docx");
    }
}
