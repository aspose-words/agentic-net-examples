using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;   // Required for the Table class

public class ExportTableAsImage
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a temporary DOCX file (required for rendering).
        string tempDocPath = Path.Combine(Path.GetTempPath(), "TempTableDoc.docx");
        doc.Save(tempDocPath, SaveFormat.Docx);

        // Reload the document to ensure all layout calculations are performed.
        Document renderedDoc = new Document(tempDocPath);

        // Configure image save options to render the first page (which contains the table) as PNG.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render only the first page where the table resides.
            PageSet = new PageSet(0)
        };

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the rendered image.
        string outputImagePath = Path.Combine(outputDir, "TableImage.png");
        renderedDoc.Save(outputImagePath, options);

        // Clean up the temporary document.
        if (File.Exists(tempDocPath))
            File.Delete(tempDocPath);
    }
}
