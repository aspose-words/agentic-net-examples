using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the document title property that will be displayed in the footer.
        doc.BuiltInDocumentProperties.Title = "My Document Title";

        // Enable different even and odd page footers.
        // In Aspose.Words the property is OddAndEvenPagesHeaderFooter.
        doc.Sections[0].PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Use DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder cursor to the even-page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

        // Align the paragraph to the left.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // Insert a field that displays the document title.
        builder.InsertField("DOCPROPERTY Title \\* MERGEFORMAT");

        // Save the document to disk.
        string outputPath = "EvenFooter.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
