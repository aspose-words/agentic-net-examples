using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeMhtmlExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source MHTML document that contains MERGEFIELDs.
            string sourceMhtml = @"C:\Docs\Template.mht";

            // Path where the merged document will be saved (can be any supported format, e.g., DOCX).
            string resultDocx = @"C:\Docs\Result.docx";

            PerformMailMergeOnMhtml(sourceMhtml, resultDocx);
        }

        /// <summary>
        /// Loads an MHTML document, performs a mail merge using a DataTable,
        /// and saves the result to the specified output file.
        /// </summary>
        /// <param name="inputMhtmlPath">Full path to the input MHTML file.</param>
        /// <param name="outputPath">Full path to the output document.</param>
        static void PerformMailMergeOnMhtml(string inputMhtmlPath, string outputPath)
        {
            // Load the MHTML document. Aspose.Words automatically detects the format.
            Document doc = new Document(inputMhtmlPath);

            // Create a DataTable that matches the merge fields present in the template.
            DataTable data = new DataTable("Customer");
            data.Columns.Add("FirstName");
            data.Columns.Add("LastName");
            data.Columns.Add("Address");
            data.Rows.Add("John", "Doe", "123 Main St, Anytown");
            data.Rows.Add("Jane", "Smith", "456 Oak Ave, Othertown");

            // Execute the mail merge. This will replace each MERGEFIELD with the corresponding data.
            // Since the document does not contain mail merge regions, the whole document will be
            // repeated for each row in the DataTable.
            doc.MailMerge.Execute(data);

            // Save the merged document. The format is inferred from the file extension.
            doc.Save(outputPath);
        }
    }
}
