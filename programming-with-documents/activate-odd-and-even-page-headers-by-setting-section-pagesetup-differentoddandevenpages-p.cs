using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsOddEvenHeaders
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Access the first (and only) section of the document.
            Section section = doc.FirstSection;

            // Enable different odd and even page headers.
            // Use reflection to set the property in case the API version does not expose it directly.
            var pageSetup = section.PageSetup;
            var prop = pageSetup.GetType().GetProperty("DifferentOddAndEvenPages");
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(pageSetup, true);
            }

            // Use DocumentBuilder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add odd page header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Odd Page Header");

            // Add even page header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Writeln("Even Page Header");

            // Return to the main body of the document.
            builder.MoveToDocumentEnd();

            // Add some text to generate at least two pages.
            builder.Writeln("This is the content of the first page.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is the content of the second page.");

            // Define output file path.
            string outputPath = "OddEvenHeaders.docx";

            // Save the document.
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
}
