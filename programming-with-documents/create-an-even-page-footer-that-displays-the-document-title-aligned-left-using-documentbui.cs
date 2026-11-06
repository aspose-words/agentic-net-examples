using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document title property.
        doc.BuiltInDocumentProperties.Title = "Sample Document Title";

        // Enable different odd and even page footers.
        // Note: In some older Aspose.Words versions the property may not be available.
        // If the property exists, uncomment the following line:
        // builder.PageSetup.DifferentOddAndEvenPages = true;

        // Move to the even page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

        // Align the paragraph to the left (default, but set explicitly).
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // Insert a field that displays the document title.
        builder.InsertField("DOCPROPERTY Title \\* MERGEFORMAT");

        // Save the document.
        doc.Save("EvenPageFooter.docx");
    }
}
