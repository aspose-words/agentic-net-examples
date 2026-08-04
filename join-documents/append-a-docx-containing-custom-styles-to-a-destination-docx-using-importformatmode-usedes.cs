using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main(string[] args)
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- Destination document ----------
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);

        // Create a custom style in the destination document.
        Style dstStyle = dstDoc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        dstStyle.Font.Color = Color.Blue; // Destination style uses blue text.

        // Apply the custom style to a paragraph.
        dstBuilder.ParagraphFormat.StyleName = "MyCustomStyle";
        dstBuilder.Writeln("Destination paragraph with custom style.");

        // ---------- Source document ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        // Create a style with the same name but different formatting.
        Style srcStyle = srcDoc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        srcStyle.Font.Color = Color.Red; // Source style uses red text.

        // Apply the custom style to a paragraph.
        srcBuilder.ParagraphFormat.StyleName = "MyCustomStyle";
        srcBuilder.Writeln("Source paragraph with custom style.");

        // ---------- Append source to destination ----------
        // Use ImportFormatMode.UseDestinationStyles to force the source content
        // to adopt the destination's style definitions.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

        // ---------- Save merged document ----------
        string mergedDocPath = Path.Combine(outputDir, "Merged.docx");
        dstDoc.Save(mergedDocPath, SaveFormat.Docx);

        // ---------- Export merged document to PDF ----------
        string pdfPath = Path.Combine(outputDir, "Merged.pdf");
        dstDoc.Save(pdfPath, SaveFormat.Pdf);

        // ---------- Validation ----------
        if (!File.Exists(mergedDocPath) || !File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Failed to create the output files.");
        }

        // Indicate successful completion (no interactive input required).
        Console.WriteLine("Documents created successfully:");
        Console.WriteLine(mergedDocPath);
        Console.WriteLine(pdfPath);
    }
}
