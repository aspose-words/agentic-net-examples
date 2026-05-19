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

        // Insert a table into the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        Table headerTable = builder.StartTable();

        // First row of the header table.
        builder.InsertCell();
        builder.Write("Header Cell 1");
        builder.InsertCell();
        builder.Write("Header Cell 2");
        builder.EndRow();

        // Second row of the header table.
        builder.InsertCell();
        builder.Write("Header Cell 3");
        builder.InsertCell();
        builder.Write("Header Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Return to the main body of the document.
        builder.MoveToSection(0);
        builder.Writeln("Body content starts here.");

        // Insert a table into the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        Table footerTable = builder.StartTable();

        // Single row footer table.
        builder.InsertCell();
        builder.Write("Footer Cell 1");
        builder.InsertCell();
        builder.Write("Footer Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderFooterTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
