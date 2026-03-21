using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create the source document with a custom style.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        Style customStyle = srcDoc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Color = System.Drawing.Color.Red;
        srcBuilder.ParagraphFormat.Style = customStyle;
        srcBuilder.Writeln("This text uses a custom style.");

        // Create the destination document.
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("This is the original destination document.");

        // Append the source document to the destination document,
        // preserving the destination's existing styles.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as PDF in the current directory.
        string outputPdfPath = "CombinedResult.pdf";
        dstDoc.Save(outputPdfPath, SaveFormat.Pdf);

        Console.WriteLine($"Combined PDF saved to: {outputPdfPath}");
    }
}
