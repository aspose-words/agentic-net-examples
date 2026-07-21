using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create the destination document with a custom style.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Define a custom paragraph style for the destination document.
        Style destStyle = destDoc.Styles.Add(StyleType.Paragraph, "CustomDestStyle");
        destStyle.Font.Name = "Arial";
        destStyle.Font.Size = 14;
        destStyle.Font.Color = System.Drawing.Color.Blue;

        // Apply the custom style and write some text.
        destBuilder.ParagraphFormat.StyleName = destStyle.Name;
        destBuilder.Writeln("This is text from the destination document.");

        // Save the destination document (optional, for inspection).
        string destPath = Path.Combine(outputDir, "Destination.docx");
        destDoc.Save(destPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the source document with its own custom style.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        // Define a custom paragraph style for the source document.
        Style srcStyle = srcDoc.Styles.Add(StyleType.Paragraph, "CustomSourceStyle");
        srcStyle.Font.Name = "Times New Roman";
        srcStyle.Font.Size = 16;
        srcStyle.Font.Color = System.Drawing.Color.DarkRed;

        // Apply the custom style and write some text.
        srcBuilder.ParagraphFormat.StyleName = srcStyle.Name;
        srcBuilder.Writeln("This is text from the source document.");

        // Save the source document (optional, for inspection).
        string srcPath = Path.Combine(outputDir, "Source.docx");
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Append the source document to the destination document using
        // ImportFormatMode.UseDestinationStyles to force the destination's
        // styles to be used for any style name clashes.
        // -----------------------------------------------------------------
        destDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

        // Save the merged document.
        string mergedDocPath = Path.Combine(outputDir, "Merged.docx");
        destDoc.Save(mergedDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Export the merged document to PDF.
        // -----------------------------------------------------------------
        string mergedPdfPath = Path.Combine(outputDir, "Merged.pdf");
        destDoc.Save(mergedPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation: ensure files exist and contain expected content.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedDocPath))
            throw new FileNotFoundException("Merged DOCX was not created.", mergedDocPath);
        if (!File.Exists(mergedPdfPath))
            throw new FileNotFoundException("Merged PDF was not created.", mergedPdfPath);

        // Verify that the merged document contains text from both source and destination.
        string mergedText = destDoc.GetText();
        if (!mergedText.Contains("This is text from the destination document.") ||
            !mergedText.Contains("This is text from the source document."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // Indicate successful completion.
        Console.WriteLine("Documents merged and exported to PDF successfully.");
    }
}
