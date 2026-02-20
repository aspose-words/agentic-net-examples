using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeMhtmlExample
{
    class Program
    {
        static void Main()
        {
            // Path to the input MHTML document that contains MERGEFIELD tags.
            string inputMhtmlPath = Path.Combine(Environment.CurrentDirectory, "Template.mht");

            // Load the MHTML document.
            Document doc = new Document(inputMhtmlPath);

            // Example data for the mail merge.
            // The field names must match the MERGEFIELD names present in the template.
            string[] fieldNames = { "FirstName", "LastName", "Email" };
            object[] fieldValues = { "John", "Doe", "john.doe@example.com" };

            // Perform the mail merge.
            // UseNonMergeFields is set to false because the template uses standard MERGEFIELD tags.
            doc.MailMerge.UseNonMergeFields = false;
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Optionally clean up any empty paragraphs that may have been left after the merge.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;

            // Save the result. Here we save as a DOCX file, but you could also save back to MHTML.
            string outputDocxPath = Path.Combine(Environment.CurrentDirectory, "MergedResult.docx");
            doc.Save(outputDocxPath);
        }
    }
}
