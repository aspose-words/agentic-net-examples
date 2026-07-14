using System;
using Aspose.Words;

namespace MailMergeFooterExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert merge fields into the main body.
            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(":");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // Add a static footer that will appear on every page.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Confidential – Company Footer Text");

            // Return the builder to the main story.
            builder.MoveToDocumentEnd();

            // Prepare data for the mail merge.
            string[] fieldNames = { "FirstName", "LastName", "Message" };
            object[] fieldValues = { "John", "Doe", "Hello! This is a merged message." };

            // Execute the mail merge.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the resulting document.
            doc.Save("MergedDocument.docx");
        }
    }
}
