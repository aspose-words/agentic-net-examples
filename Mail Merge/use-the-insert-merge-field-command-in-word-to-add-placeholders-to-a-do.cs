using System;
using Aspose.Words;

namespace InsertMergeFieldsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a merge field for "FirstName".
            // The field code is inserted without the surrounding curly braces.
            builder.InsertField("MERGEFIELD FirstName \\* MERGEFORMAT");

            // Insert a space between fields.
            builder.Write(" ");

            // Insert a merge field for "LastName".
            builder.InsertField("MERGEFIELD LastName \\* MERGEFORMAT");

            // Insert a paragraph break.
            builder.Writeln();

            // Insert a merge field for "Address".
            builder.InsertField("MERGEFIELD Address \\* MERGEFORMAT");

            // Save the document to a DOCX file.
            doc.Save("MergeFieldsDocument.docx");
        }
    }
}
