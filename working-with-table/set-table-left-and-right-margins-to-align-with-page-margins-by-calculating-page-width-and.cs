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
        // Access the first (and only) section to configure page setup.
        Section section = doc.FirstSection;
        PageSetup pageSetup = section.PageSetup;

        // Set custom page margins (optional, just for demonstration).
        pageSetup.LeftMargin = 72;   // 1 inch
        pageSetup.RightMargin = 72;  // 1 inch
        pageSetup.TopMargin = 72;
        pageSetup.BottomMargin = 72;

        // Build a simple 2‑column table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Column 1");
        builder.InsertCell();
        builder.Write("Column 2");
        builder.EndRow();
        builder.EndTable();

        // Calculate the usable page width (page width minus left and right margins).
        double usablePageWidth = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;

        // Align the table with the page margins.
        // LeftIndent moves the table away from the left page edge.
        table.LeftIndent = pageSetup.LeftMargin;
        // PreferredWidth defines the total width of the table.
        table.PreferredWidth = PreferredWidth.FromPoints(usablePageWidth);

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableAlignedWithMargins.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
