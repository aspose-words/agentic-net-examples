using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC field at the beginning of the document.
        // The \b switch tells the TOC to include entries from the specified bookmark.
        builder.InsertField(@"TOC \b Table1");
        builder.Writeln(); // Add a paragraph break after the TOC.

        // Insert a bookmark that will be referenced by the TOC.
        builder.StartBookmark("Table1");
        builder.Writeln("Table 1: Sample Table");
        builder.EndBookmark("Table1");

        // Build a simple 2x2 table after the bookmark.
        builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row – data cells.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Update fields so the TOC reflects the bookmark entry.
        doc.UpdateFields();

        // Save the document to the output file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
