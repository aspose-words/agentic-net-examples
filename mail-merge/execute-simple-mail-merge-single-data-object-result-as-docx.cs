using System;
using Aspose.Words;

namespace MailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert merge fields that will be populated during the mail merge.
            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(":");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // Prepare the data for a single record.
            string[] fieldNames = { "FirstName", "LastName", "Message" };
            object[] fieldValues = { "John", "Doe", "Hello! This message was created with Aspose.Words mail merge." };

            // Execute the mail merge for the single record.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the merged document as DOCX.
            doc.Save("SimpleMailMergeResult.docx");
        }
    }
}
