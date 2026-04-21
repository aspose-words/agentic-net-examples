using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableMarginAlignment
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();
            builder.EndTable();

            // Retrieve page setup information.
            PageSetup pageSetup = doc.FirstSection.PageSetup;
            double leftMargin = pageSetup.LeftMargin;   // points
            double rightMargin = pageSetup.RightMargin; // points
            double pageWidth = pageSetup.PageWidth;     // total page width in points

            // Calculate the usable width between the margins.
            double usableWidth = pageWidth - leftMargin - rightMargin;

            // Align the table with the page margins.
            table.LeftIndent = leftMargin;
            table.PreferredWidth = PreferredWidth.FromPoints(usableWidth);

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlignedTable.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
