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

        // Move the builder cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Write the static text "Page ".
        builder.Write("Page ");

        // Insert a PAGE field with a switch to enforce Arabic numerals.
        // The second parameter is an empty placeholder for the field result.
        builder.InsertField("PAGE \\* Arabic", "");

        // Write the static text " of ".
        builder.Write(" of ");

        // Insert a NUMPAGES field with the same Arabic numeral switch.
        builder.InsertField("NUMPAGES \\* Arabic", "");

        // Update all fields so they display the correct values.
        doc.UpdateFields();

        // Save the document to a file in the current directory.
        const string outputPath = "PageNumberFooter.docx";
        doc.Save(outputPath);
    }
}
