using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the intermediate DOCX and the final PDF.
        const string docxPath = "SampleTable.docx";
        const string pdfPath = "SampleTable.pdf";

        // -------------------------------------------------
        // Step 1: Create a DOCX document that contains a formatted table.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title paragraph.
        builder.Writeln("Document with a Sample Table");
        builder.Writeln();

        // Begin the table.
        builder.StartTable();

        // ----- Header row -----
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Writeln("Header 1");

        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Writeln("Header 2");

        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Writeln("Header 3");
        builder.EndRow();

        // ----- First data row -----
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Writeln("Row 1, Cell 1");

        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");

        builder.InsertCell();
        builder.Writeln("Row 1, Cell 3");
        builder.EndRow();

        // ----- Second data row -----
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");

        builder.InsertCell();
        builder.Writeln("Row 2, Cell 2");

        builder.InsertCell();
        builder.Writeln("Row 2, Cell 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table and apply formatting.
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
            throw new InvalidOperationException("Failed to locate the created table.");

        table.StyleIdentifier = StyleIdentifier.TableGrid;
        table.AllowAutoFit = true;

        // Save the DOCX to disk.
        doc.Save(docxPath);

        // -------------------------------------------------
        // Step 2: Load the DOCX containing the table.
        // -------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // -------------------------------------------------
        // Step 3: Export the loaded document to PDF while preserving table styles.
        // -------------------------------------------------
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // Validation: Ensure the PDF file was created successfully.
        // -------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create PDF file at '{pdfPath}'.");

        // Optional cleanup of the intermediate DOCX.
        // File.Delete(docxPath);
    }
}
