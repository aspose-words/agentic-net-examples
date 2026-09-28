using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files
        string destDocPath = "Destination.docx";
        string sourceDocPath = "Source.docx";
        string mergedPdfPath = "Merged.pdf";

        // -----------------------------------------------------------------
        // Create destination document with a custom style
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        Style destStyle = destDoc.Styles.Add(StyleType.Paragraph, "CustomStyleDest");
        destStyle.Font.Name = "Courier New";
        destStyle.Font.Size = 14;
        destStyle.Font.Color = System.Drawing.Color.DarkBlue;

        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.ParagraphFormat.Style = destStyle;
        destBuilder.Writeln("This is a paragraph in the destination document using CustomStyleDest.");
        destDoc.Save(destDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create source document with its own custom style
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        Style sourceStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "CustomStyleSource");
        sourceStyle.Font.Name = "Times New Roman";
        sourceStyle.Font.Size = 12;
        sourceStyle.Font.Color = System.Drawing.Color.Maroon;

        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.ParagraphFormat.Style = sourceStyle;
        sourceBuilder.Writeln("This is a paragraph in the source document using CustomStyleSource.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents (optional, they are already in memory)
        // -----------------------------------------------------------------
        Document destination = new Document(destDocPath);
        Document source = new Document(sourceDocPath);

        // Append source to destination using destination styles
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the merged document as PDF
        destination.Save(mergedPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPdfPath))
        {
            throw new InvalidOperationException($"The merged PDF file was not created at '{mergedPdfPath}'.");
        }

        // Verify that the merged document contains sections from both source documents
        if (destination.Sections.Count < 2)
        {
            throw new InvalidOperationException("The merged document does not contain the expected number of sections.");
        }

        // Program completed successfully
    }
}
