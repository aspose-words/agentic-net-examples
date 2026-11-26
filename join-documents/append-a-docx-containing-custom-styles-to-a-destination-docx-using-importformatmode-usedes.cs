using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string destDocPath = "Destination.docx";
        string srcDocPath = "Source.docx";
        string outputPdfPath = "MergedOutput.pdf";

        // ---------- Create destination document with a custom style ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Define a custom paragraph style for the destination document
        Style destStyle = destDoc.Styles.Add(StyleType.Paragraph, "DestCustomStyle");
        destStyle.Font.Color = Color.DarkGreen;
        destStyle.Font.Size = 16;
        destStyle.Font.Bold = true;

        // Apply the custom style and add some content
        destBuilder.ParagraphFormat.Style = destStyle;
        destBuilder.Writeln("This is a paragraph in the destination document using DestCustomStyle.");

        // Save the destination document (required before appending)
        destDoc.Save(destDocPath, SaveFormat.Docx);

        // ---------- Create source document with its own custom style ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        // Define a custom paragraph style for the source document
        Style srcStyle = srcDoc.Styles.Add(StyleType.Paragraph, "SrcCustomStyle");
        srcStyle.Font.Color = Color.DarkRed;
        srcStyle.Font.Size = 14;
        srcStyle.Font.Italic = true;

        // Apply the custom style and add some content
        srcBuilder.ParagraphFormat.Style = srcStyle;
        srcBuilder.Writeln("This is a paragraph in the source document using SrcCustomStyle.");

        // Save the source document
        srcDoc.Save(srcDocPath, SaveFormat.Docx);

        // ---------- Load documents (simulating typical usage) ----------
        Document destination = new Document(destDocPath);
        Document source = new Document(srcDocPath);

        // Append source document to destination using destination styles
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the merged document as PDF
        destination.Save(outputPdfPath, SaveFormat.Pdf);

        // ---------- Validation ----------
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException($"Failed to create PDF output at '{outputPdfPath}'.");
        }

        string mergedText = destination.Range.Text;
        if (!mergedText.Contains("This is a paragraph in the destination document") ||
            !mergedText.Contains("This is a paragraph in the source document"))
        {
            throw new InvalidOperationException("Merged document does not contain expected content from both source files.");
        }

        // Clean up temporary DOCX files (optional)
        File.Delete(destDocPath);
        File.Delete(srcDocPath);
    }
}
