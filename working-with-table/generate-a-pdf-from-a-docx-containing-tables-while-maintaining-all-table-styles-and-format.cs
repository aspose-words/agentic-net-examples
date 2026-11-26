using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary DOCX and final PDF files.
        string docxPath = "SampleTable.docx";
        string pdfPath = "SampleTable.pdf";

        // -----------------------------------------------------------------
        // Create a DOCX document with a formatted table.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // ---- First row (header) ----
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // ---- Second row (data) ----
        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("10");
        builder.InsertCell();
        builder.Writeln("$1.20");
        builder.EndRow();

        // ---- Third row (data) ----
        builder.InsertCell();
        builder.Writeln("Bananas");
        builder.InsertCell();
        builder.Writeln("5");
        builder.InsertCell();
        builder.Writeln("$0.80");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Apply a built‑in table style.
        Table table = doc.FirstSection.Body.Tables[0];
        table.StyleIdentifier = StyleIdentifier.TableGrid;

        // Apply shading to the header row cells (RowFormat.Shading is not available in this version).
        Row headerRow = table.Rows[0];
        foreach (Cell cell in headerRow.Cells)
        {
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        }

        // Save the DOCX to disk.
        doc.Save(docxPath);

        // -----------------------------------------------------------------
        // Load the DOCX and convert it to PDF while preserving table styles.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new Exception("PDF conversion failed: output file not found.");
        }

        // Optional cleanup of the temporary DOCX.
        // File.Delete(docxPath);
    }
}
