using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert merge fields that will be filled during the mail merge.
            builder.InsertField(" MERGEFIELD FullName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Company ");

            // Define the field names and corresponding values for a single record.
            string[] fieldNames = { "FullName", "Company" };
            object[] fieldValues = { "James Bond", "MI5 Headquarters" };

            // Perform the mail merge.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the merged document as a DOCX file.
            doc.Save("MergedDocument.docx", SaveFormat.Docx);
        }
    }
}
