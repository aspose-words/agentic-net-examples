using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some content to have multiple pages.
            builder.Writeln("First page content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Second page content.");

            // Move the builder to the end of the document to insert the paragraph with the page number field.
            builder.MoveToDocumentEnd();

            // Insert a paragraph that contains the page number field.
            builder.Write("Page ");
            // Insert a PAGE field; the result will be updated later.
            builder.InsertField("PAGE", "");
            builder.Writeln();

            // Configure the section to display page numbers as uppercase Roman numerals.
            PageSetup pageSetup = doc.FirstSection.PageSetup;
            pageSetup.RestartPageNumbering = true;      // Start numbering from the beginning of the section.
            pageSetup.PageStartingNumber = 1;           // First page number.
            pageSetup.PageNumberStyle = NumberStyle.UppercaseRoman; // Roman numeral format.

            // Update all fields in the document so the PAGE field shows the correct value.
            doc.UpdateFields();

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "RomanPageNumbers.docx");
            doc.Save(outputPath);
        }
    }
}
