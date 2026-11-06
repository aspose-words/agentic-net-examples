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

        // Insert a custom TOC field that includes only entries with the "Table" caption label.
        builder.InsertField(@"TOC \h \z \c ""Table""");
        builder.Writeln(); // Move to a new line after the TOC.
        builder.InsertBreak(BreakType.PageBreak); // Start tables on a new page.

        // ---------- First table ----------
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Insert caption for the first table (manual implementation of InsertCaption).
        // SEQ field for the table number.
        builder.InsertField("SEQ Table \\* ARABIC");
        // Write the caption text after the SEQ field.
        builder.Write(" Sample Table 1");
        // TC field so that the custom TOC can pick up this entry.
        builder.InsertField(@"TC ""Table"" \f t");
        builder.Writeln(); // Add spacing after caption.

        // ---------- Second table ----------
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Column A");
        builder.InsertCell();
        builder.Writeln("Column B");
        builder.InsertCell();
        builder.Writeln("Column C");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Data A1");
        builder.InsertCell();
        builder.Writeln("Data B1");
        builder.InsertCell();
        builder.Writeln("Data C1");
        builder.EndRow();
        builder.EndTable();

        // Insert caption for the second table (manual implementation of InsertCaption).
        builder.InsertField("SEQ Table \\* ARABIC");
        builder.Write(" Sample Table 2");
        builder.InsertField(@"TC ""Table"" \f t");
        builder.Writeln();

        // Update fields (TOC and SEQ fields) to reflect the inserted captions.
        doc.UpdateFields();

        // Save the document.
        string outputPath = "TableOfContents.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
