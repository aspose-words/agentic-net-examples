using System;
using Aspose.Words;

namespace MailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Path to the folder that contains the template DOCX file.
            // Adjust this path to point to your actual location.
            string docsPath = @"C:\Docs\";

            // Load the source document that contains MERGEFIELDs.
            Document doc = new Document(docsPath + "Template.docx");

            // Define the merge field names present in the template and the corresponding values.
            string[] fieldNames = { "FullName", "Address" };
            object[] fieldValues = { "John Doe", "123 Main St., Anytown" };

            // Execute a standard mail merge for a single record.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the merged document.
            doc.Save(docsPath + "MergedOutput.docx");
        }
    }
}
