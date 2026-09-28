using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some merge fields into the body of the document.
        builder.Writeln("Dear ");

        // Insert MERGEFIELD Name.
        builder.InsertField(FieldType.FieldMergeField, true);
        builder.Write(" Name");

        builder.Writeln(",");

        builder.Writeln("Your order number is ");

        // Insert MERGEFIELD OrderNumber.
        builder.InsertField(FieldType.FieldMergeField, true);
        builder.Write(" OrderNumber");

        builder.Writeln(".");

        // Add static footer text using DocumentBuilder.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("This is a static footer added before mail merge.");

        // Prepare data for mail merge.
        string[] fieldNames = { "Name", "OrderNumber" };
        object[] fieldValues = { "John Doe", "12345" };

        // Execute mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the result to a file.
        doc.Save("MergedDocument.docx");
    }
}
