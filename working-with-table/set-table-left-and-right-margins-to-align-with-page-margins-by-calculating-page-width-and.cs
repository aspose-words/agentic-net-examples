using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableMarginAlignmentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Set custom page margins (optional, just to demonstrate the calculation).
            // Values are in points (1 inch = 72 points).
            doc.FirstSection.PageSetup.LeftMargin = 72;   // 1 inch
            doc.FirstSection.PageSetup.RightMargin = 72;  // 1 inch
            doc.FirstSection.PageSetup.TopMargin = 72;
            doc.FirstSection.PageSetup.BottomMargin = 72;

            // Build a simple 2‑column table.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Calculate the usable page width (page width minus left and right margins).
            PageSetup pageSetup = doc.FirstSection.PageSetup;
            double usablePageWidth = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;

            // Align the table with the page margins.
            // Set the left indent of the table to match the left page margin.
            table.LeftIndent = pageSetup.LeftMargin;

            // Set the preferred width of the table to the usable page width.
            // This makes the right edge of the table line up with the right page margin.
            table.PreferredWidth = PreferredWidth.FromPoints(usablePageWidth);

            // Save the document.
            string outputPath = "TableAlignedWithMargins.docx";
            doc.Save(outputPath);
        }
    }
}
