using System;
using Aspose.Words;

namespace MailMergeFooterExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert merge fields into the main body.
            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(".");

            // Add static text to the footer of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Confidential - Company Footer");

            // Data for the mail merge.
            string[] fieldNames = { "FirstName", "LastName" };
            object[] fieldValues = { "John", "Doe" };

            // Perform the mail merge.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the merged document.
            doc.Save("MergedDocument.docx");
        }
    }
}
