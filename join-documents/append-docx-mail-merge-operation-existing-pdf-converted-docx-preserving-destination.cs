using System;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Paths to the files.
        string existingDocPath = "ExistingFromPdf.docx";   // DOCX that originated from a PDF.
        string templatePath    = "MailMergeTemplate.docx"; // Template containing MERGEFIELDs.
        string outputPath      = "CombinedResult.docx";    // Resulting combined document.

        // Load the existing DOCX (converted from PDF) – destination document.
        Document destination = new Document(existingDocPath);

        // Load the mail‑merge template – source document.
        Document source = new Document(templatePath);

        // -------------------------------------------------
        // Prepare a simple data source for the mail merge.
        // -------------------------------------------------
        DataTable data = new DataTable("Data");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");
        data.Rows.Add("John", "Doe", "Hello from mail merge!");
        data.Rows.Add("Jane", "Smith", "Another message.");

        // Use the first row as a single‑record data source.
        string[] fieldNames = data.Columns.Cast<DataColumn>()
                                          .Select(col => col.ColumnName)
                                          .ToArray();
        object[] fieldValues = data.Rows[0].ItemArray;

        // Execute the mail merge on the source document.
        source.MailMerge.Execute(fieldNames, fieldValues);

        // -------------------------------------------------
        // Append the mail‑merged document to the destination,
        // preserving the destination's styles.
        // -------------------------------------------------
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the combined document.
        destination.Save(outputPath);
    }
}
