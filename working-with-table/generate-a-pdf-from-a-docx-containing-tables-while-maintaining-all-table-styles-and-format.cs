using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableToPdf
{
    public class Program
    {
        public static void Main()
        {
            // Define output folder and ensure it exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for intermediate DOCX and final PDF files.
            string docxPath = Path.Combine(outputDir, "TableDocument.docx");
            string pdfPath = Path.Combine(outputDir, "TableDocument.pdf");

            // -------------------------------------------------
            // 1. Create a new blank document.
            // -------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------------------------------------
            // 2. Build a table with a built‑in style.
            // -------------------------------------------------
            // Start the table.
            Table table = builder.StartTable();

            // First row – header cells.
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Second row – data cells.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("30");
            builder.EndRow();

            // Third row – data cells.
            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("45");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply a built‑in table style and enable row banding.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

            // -------------------------------------------------
            // 3. Convert style formatting to direct formatting.
            //    This ensures the visual appearance is preserved
            //    when saving to formats that do not support table styles.
            // -------------------------------------------------
            doc.ExpandTableStylesToDirectFormatting();

            // -------------------------------------------------
            // 4. Save the document as DOCX (optional, shows the source file).
            // -------------------------------------------------
            doc.Save(docxPath);

            // -------------------------------------------------
            // 5. Save the same document as PDF.
            // -------------------------------------------------
            doc.Save(pdfPath, SaveFormat.Pdf);

            // -------------------------------------------------
            // 6. Indicate completion.
            // -------------------------------------------------
            Console.WriteLine($"DOCX saved to: {docxPath}");
            Console.WriteLine($"PDF saved to: {pdfPath}");
        }
    }
}
