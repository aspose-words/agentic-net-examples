using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the temporary documents and final PDF.
        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.docx");
        string mergedPdfPath = Path.Combine(outputDir, "Merged.pdf");

        // ---------- Create destination document ----------
        var destDoc = new Document();
        var destBuilder = new DocumentBuilder(destDoc);

        // Define a custom style in the destination document.
        Style destCustomStyle = destDoc.Styles.Add(StyleType.Paragraph, "CustomStyle");
        destCustomStyle.Font.Name = "Arial";
        destCustomStyle.Font.Size = 16;
        destCustomStyle.Font.Color = Color.Blue;

        // Apply the custom style and add some text.
        destBuilder.ParagraphFormat.StyleName = "CustomStyle";
        destBuilder.Writeln("Destination document with custom style.");

        // Save the destination document as DOCX.
        destDoc.Save(destPath, SaveFormat.Docx);

        // ---------- Create source document ----------
        var srcDoc = new Document();
        var srcBuilder = new DocumentBuilder(srcDoc);

        // Define a style with the same name but different formatting.
        Style srcCustomStyle = srcDoc.Styles.Add(StyleType.Paragraph, "CustomStyle");
        srcCustomStyle.Font.Name = "Times New Roman";
        srcCustomStyle.Font.Size = 14;
        srcCustomStyle.Font.Color = Color.Red;

        // Apply the custom style and add some text.
        srcBuilder.ParagraphFormat.StyleName = "CustomStyle";
        srcBuilder.Writeln("Source document with custom style.");

        // Save the source document as DOCX.
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // ---------- Append source to destination using destination styles ----------
        // Load the documents (optional, we already have them in memory).
        var destination = new Document(destPath);
        var source = new Document(srcPath);

        // Append while preserving destination styles.
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Verify that both pieces of text are present after the append.
        string mergedText = destination.GetText();
        if (!mergedText.Contains("Destination document with custom style.") ||
            !mergedText.Contains("Source document with custom style."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // ---------- Save the merged document as PDF ----------
        destination.Save(mergedPdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(mergedPdfPath))
        {
            throw new FileNotFoundException("The merged PDF was not created.", mergedPdfPath);
        }

        // Optional: indicate success.
        Console.WriteLine("Documents merged and saved to PDF successfully.");
    }
}
