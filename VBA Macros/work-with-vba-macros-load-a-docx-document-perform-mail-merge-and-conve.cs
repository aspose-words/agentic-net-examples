using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX/DOCM file that contains MERGEFIELD tags.
            const string sourcePath = @"C:\Docs\TemplateWithMacros.docx";

            // Path where the resulting PDF will be saved.
            const string pdfOutputPath = @"C:\Docs\ResultReport.pdf";

            // Load the existing document (lifecycle rule: load).
            Document doc = new Document(sourcePath);

            // Prepare a simple data source for the mail merge.
            // The column names must match the MERGEFIELD names in the template.
            DataTable mergeData = new DataTable("MergeData");
            mergeData.Columns.Add("FullName");
            mergeData.Columns.Add("Company");
            mergeData.Columns.Add("Address");
            mergeData.Columns.Add("City");

            // Add a single record – you can add more rows if you need multiple documents.
            mergeData.Rows.Add("James Bond", "MI5 Headquarters", "Milbank", "London");

            // Execute the mail merge (feature rule: MailMerge.Execute(string[], object[])).
            // Here we use the overload that accepts a DataTable for convenience.
            doc.MailMerge.Execute(mergeData);

            // Convert the merged document to PDF (lifecycle rule: save with format).
            doc.Save(pdfOutputPath, SaveFormat.Pdf);
        }
    }
}
