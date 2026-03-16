using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeMultipleDocuments
{
    class Program
    {
        static void Main()
        {
            // Load the mail‑merge template. The template must contain MERGEFIELDs that match the column names.
            Document template = new Document("Template.docx");

            // -----------------------------------------------------------------
            // Prepare a data source. In this example we use a DataTable with three rows.
            // -----------------------------------------------------------------
            DataTable data = new DataTable("Customers");
            data.Columns.Add("FirstName");
            data.Columns.Add("LastName");
            data.Columns.Add("Address");

            data.Rows.Add("John", "Doe", "120 Hanover Sq., London");
            data.Rows.Add("Jane", "Smith", "15 Main St., New York");
            data.Rows.Add("Bob", "Johnson", "42 Oak Ave., Toronto");

            // -----------------------------------------------------------------
            // Create a destination document that will hold all merged documents.
            // -----------------------------------------------------------------
            Document mergedAll = new Document();

            // Ensure the destination has at least one section (required by Aspose.Words).
            mergedAll.FirstSection.Body.AppendParagraph("Merged Documents:");

            // -----------------------------------------------------------------
            // For each data row: clone the template, perform a single‑record mail merge,
            // and append the result to the destination document.
            // -----------------------------------------------------------------
            foreach (DataRow row in data.Rows)
            {
                // Clone the template. The true argument performs a deep clone (including styles).
                Document singleMerged = (Document)template.Clone(true);

                // Execute a mail merge for the current row only.
                singleMerged.MailMerge.Execute(row);

                // Append the merged document to the master document.
                // KeepSourceFormatting preserves the original formatting of the cloned document.
                mergedAll.AppendDocument(singleMerged, ImportFormatMode.KeepSourceFormatting);

                // Optional: insert a page break between merged documents for readability.
                DocumentBuilder builder = new DocumentBuilder(mergedAll);
                builder.MoveToDocumentEnd();
                builder.InsertBreak(BreakType.PageBreak);
            }

            // -----------------------------------------------------------------
            // Save the combined document that contains all merged outputs.
            // -----------------------------------------------------------------
            mergedAll.Save("AllMergedDocuments.docx", SaveFormat.Docx);
        }
    }
}
