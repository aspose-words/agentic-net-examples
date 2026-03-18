using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace MergeFieldExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a MERGEFIELD for the customer's name.
            // The field code must be written without the surrounding braces.
            builder.InsertField(" MERGEFIELD CustomerName ");

            // Insert a paragraph break between the fields.
            builder.InsertParagraph();

            // Insert a MERGEFIELD for the customer's address.
            builder.InsertField(" MERGEFIELD Address ");

            // Save the document to disk.
            // Adjust the path as needed for your environment.
            doc.Save("CustomerTemplate.docx");
        }
    }
}
