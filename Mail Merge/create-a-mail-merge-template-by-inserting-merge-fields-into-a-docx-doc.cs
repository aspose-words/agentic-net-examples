using System;
using Aspose.Words;

class MailMergeTemplateCreator
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert merge fields into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a greeting line with two merge fields: FirstName and LastName.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln(":");

        // Insert an address block with an Address merge field.
        builder.InsertField("MERGEFIELD Address");

        // Save the document as a DOCX file.
        doc.Save("MailMergeTemplate.docx");
    }
}
