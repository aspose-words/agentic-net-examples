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

        // -------------------------------------------------
        // Insert a table into the primary header.
        // -------------------------------------------------
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

        // Finish the table.
        builder.EndTable();

        // -------------------------------------------------
        // Insert a table into the primary footer.
        // -------------------------------------------------
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        Table footerTable = builder.StartTable();

        // Single row in the footer.
        builder.InsertCell();
        builder.Write("Footer Cell 1");
        builder.InsertCell();
        builder.Write("Footer Cell 2");
        builder.EndRow();

        builder.EndTable();

        // -------------------------------------------------
        // Add some regular body content so the document is not empty.
        // -------------------------------------------------
        builder.MoveToSection(0);
        builder.Writeln("This is the main body of the document.");

        // -------------------------------------------------
        // Save the document.
        // -------------------------------------------------
        string outputPath = Path.Combine(Environment.CurrentDirectory, "HeaderFooterTable.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
