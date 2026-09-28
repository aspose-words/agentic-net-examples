using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a greeting line with a merge field for the customer's name.
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD CustomerName \\* MERGEFORMAT");
        builder.Writeln(",");

        // Insert a line with a merge field for the customer's address.
        builder.Writeln("Your address is:");
        builder.InsertField("MERGEFIELD CustomerAddress \\* MERGEFORMAT");
        builder.Writeln(".");

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
