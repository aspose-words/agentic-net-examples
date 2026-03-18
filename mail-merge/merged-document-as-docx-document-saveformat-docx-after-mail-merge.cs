using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert merge fields that will be populated.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertField(" MERGEFIELD FullName ");
        builder.InsertParagraph();
        builder.InsertField(" MERGEFIELD Company ");

        // Perform a simple mail merge with sample data.
        string[] fieldNames = { "FullName", "Company" };
        object[] fieldValues = { "James Bond", "MI5 Headquarters" };
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged result as a DOCX file.
        doc.Save("MergedDocument.docx", SaveFormat.Docx);
    }
}
