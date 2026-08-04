using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert merge fields into the document.
        builder.InsertField("MERGEFIELD FullName");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Company");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Address");
        builder.Writeln();
        builder.InsertField("MERGEFIELD City");
        builder.Writeln();

        // Prepare data for a single record mail merge.
        string[] fieldNames = { "FullName", "Company", "Address", "City" };
        object[] fieldValues = { "James Bond", "MI5 Headquarters", "Milbank", "London" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Define the output PDF file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedDocument.pdf");

        // Save the merged document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
