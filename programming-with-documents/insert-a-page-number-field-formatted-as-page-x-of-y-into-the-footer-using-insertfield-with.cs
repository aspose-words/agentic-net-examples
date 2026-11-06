using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert the "Page X of Y" field.
        // Current page number.
        builder.InsertField("PAGE \\* MERGEFORMAT", "1");
        // Static text separator.
        builder.Write(" of ");
        // Total number of pages.
        builder.InsertField("NUMPAGES \\* MERGEFORMAT", "1");

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
