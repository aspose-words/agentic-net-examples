using System;
using System.IO;
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

            // Access the first (and only) section to retrieve page setup information.
            PageSetup pageSetup = doc.FirstSection.PageSetup;

            // Build a simple 2‑column table with a few rows.
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
            builder.Write("Cell A1");
            builder.InsertCell();
            builder.Write("Cell A2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Align the table with the page margins.
            // Left indent = left page margin.
            table.LeftIndent = pageSetup.LeftMargin;

            // Preferred width = page width minus left and right margins.
            double pageWidth = pageSetup.PageWidth;
            double leftMargin = pageSetup.LeftMargin;
            double rightMargin = pageSetup.RightMargin;
            table.PreferredWidth = PreferredWidth.FromPoints(pageWidth - leftMargin - rightMargin);

            // Save the document.
            string outputPath = "TableMarginsAligned.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
