using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Enable different odd and even page headers/footers for the first section.
        // The property DifferentOddAndEvenPages may not be available in older library versions,
        // so we set it via reflection to keep the code compilable across versions.
        Section section = doc.FirstSection;
        var pageSetup = section.PageSetup;
        var prop = pageSetup.GetType().GetProperty("DifferentOddAndEvenPages");
        if (prop != null && prop.CanWrite)
        {
            prop.SetValue(pageSetup, true);
        }

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an odd page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Odd Page Header");

        // Add an even page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Writeln("Even Page Header");

        // Add enough body text to generate multiple pages.
        builder.MoveToDocumentEnd();
        for (int i = 1; i <= 30; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Save the document.
        string outputPath = "OddEvenHeaders.docx";
        doc.Save(outputPath);

        // Reload the document to verify the setting (if the property exists).
        Document loadedDoc = new Document(outputPath);
        bool different = false;
        var loadedProp = loadedDoc.FirstSection.PageSetup.GetType()
            .GetProperty("DifferentOddAndEvenPages");
        if (loadedProp != null && loadedProp.CanRead)
        {
            different = (bool)loadedProp.GetValue(loadedDoc.FirstSection.PageSetup);
        }

        Console.WriteLine($"DifferentOddAndEvenPages set: {different}");
    }
}
