using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a page number field formatted as Roman numerals.
        builder.Writeln("Page number in Roman numerals:");
        // Insert a PAGE field with the Roman numeral format switch.
        builder.InsertField("PAGE  \\* ROMAN");

        // Save the document to a file.
        const string outputPath = "RomanPageNumber.docx";
        doc.Save(outputPath);
    }
}
