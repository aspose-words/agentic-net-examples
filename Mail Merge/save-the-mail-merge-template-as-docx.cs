using System;
using Aspose.Words;

namespace MailMergeTemplateExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add mail merge fields to the template.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a few sample merge fields.
            builder.InsertField(" MERGEFIELD FullName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Company ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Address ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD City ");

            // Save the mail merge template as a DOCX file.
            // The Save(string) overload automatically determines the format from the file extension.
            doc.Save("MailMergeTemplate.docx");
        }
    }
}
