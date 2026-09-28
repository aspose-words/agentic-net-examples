using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert the "Page X of Y" text using fields with switches.
        // PAGE field with MERGEFORMAT switch.
        builder.InsertField("PAGE  \\* MERGEFORMAT", "1");
        // Static separator text.
        builder.Write(" of ");
        // NUMPAGES field with MERGEFORMAT switch.
        builder.InsertField("NUMPAGES  \\* MERGEFORMAT", "1");

        // Save the document to a file.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
