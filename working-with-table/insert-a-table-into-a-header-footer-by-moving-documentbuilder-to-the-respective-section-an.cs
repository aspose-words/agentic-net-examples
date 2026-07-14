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

        // Enable different first page header/footer (optional, shows usage of multiple headers/footers).
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // ----- Insert a table into the primary header -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        // Start the table.
        Table headerTable = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Header Cell 1");
        builder.InsertCell();
        builder.Write("Header Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Header Cell 3");
        builder.InsertCell();
        builder.Write("Header Cell 4");
        builder.EndRow();

        // Finish the header table.
        builder.EndTable();

        // ----- Insert a table into the primary footer -----
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        Table footerTable = builder.StartTable();

        // Single row with three cells.
        builder.InsertCell();
        builder.Write("Footer Col 1");
        builder.InsertCell();
        builder.Write("Footer Col 2");
        builder.InsertCell();
        builder.Write("Footer Col 3");
        builder.EndRow();

        // Finish the footer table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "HeaderFooterTable.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate successful completion.
        Console.WriteLine("Document saved to: " + Path.GetFullPath(outputPath));
    }
}
