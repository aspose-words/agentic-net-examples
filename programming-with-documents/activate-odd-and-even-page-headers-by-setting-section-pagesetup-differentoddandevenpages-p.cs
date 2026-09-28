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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different odd and even page headers/footers for the first section.
        // Note: In some older Aspose.Words versions the property may not be available.
        // If the property exists, uncomment the line below.
        // doc.FirstSection.PageSetup.DifferentOddAndEvenPages = true;

        // Add content to the odd (primary) header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Odd Page Header");

        // Add content to the even header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Writeln("Even Page Header");

        // Add some body text to generate multiple pages.
        builder.MoveToDocumentEnd();
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"Paragraph {i + 1}: Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document.
        string outputPath = "OddEvenHeaders.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
