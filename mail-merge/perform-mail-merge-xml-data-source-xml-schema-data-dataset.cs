using System;
using Aspose.Words;

class MailMergeXmlExample
{
    static void Main()
    {
        // Create a new document with simple mail‑merge fields.
        Document doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD", "FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD", "LastName");
        builder.Writeln(",");

        // Prepare data for the merge.
        string[] fieldNames = { "FirstName", "LastName" };
        object[] fieldValues = { "John", "Doe" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        doc.Save("MergedOutput.docx");
    }
}
