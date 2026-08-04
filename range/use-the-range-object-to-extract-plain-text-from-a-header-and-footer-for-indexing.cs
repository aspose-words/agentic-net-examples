using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace HeaderFooterRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Define file paths.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);
            string docPath = Path.Combine(outputDir, "SampleDocument.docx");
            string indexPath = Path.Combine(outputDir, "HeaderFooterIndex.txt");

            // -----------------------------------------------------------------
            // 1. Create a new document and add header, footer, and body text.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a primary header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header text for indexing");

            // Add a primary footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Footer text for indexing");

            // Return to the main body and add some content.
            builder.MoveToDocumentEnd();
            builder.Writeln("This is the body of the document.");

            // Save the document.
            doc.Save(docPath);

            // -----------------------------------------------------------------
            // 2. Extract plain text from the header and footer using Range.
            // -----------------------------------------------------------------
            // Load the document (demonstrates loading workflow).
            Document loadedDoc = new Document(docPath);

            // Retrieve the primary header and footer.
            HeaderFooter header = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            HeaderFooter footer = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            // Use the Range.Text property to get plain text.
            string headerText = header?.Range?.Text?.Trim() ?? string.Empty;
            string footerText = footer?.Range?.Text?.Trim() ?? string.Empty;

            // -----------------------------------------------------------------
            // 3. Write the extracted texts to a plain‑text file for indexing.
            // -----------------------------------------------------------------
            using (StreamWriter writer = new StreamWriter(indexPath))
            {
                writer.WriteLine("Header:");
                writer.WriteLine(headerText);
                writer.WriteLine();
                writer.WriteLine("Footer:");
                writer.WriteLine(footerText);
            }

            // Output the extracted texts to the console (non‑interactive).
            Console.WriteLine("Extracted Header Text: " + headerText);
            Console.WriteLine("Extracted Footer Text: " + footerText);
        }
    }
}
