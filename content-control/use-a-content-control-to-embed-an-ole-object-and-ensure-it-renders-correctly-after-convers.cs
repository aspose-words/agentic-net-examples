using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a simple text file that will be embedded as an OLE object.
        string oleSourcePath = Path.Combine(outputDir, "sample.txt");
        File.WriteAllText(oleSourcePath, "This is a sample OLE embedded file.");

        // Create a new Word document.
        Document doc = new Document();

        // Create a block‑level rich‑text content control to host the OLE object.
        StructuredDocumentTag oleSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "OleObjectControl",
            Tag = "ole-object"
        };

        // Append the content control to the document body.
        doc.FirstSection.Body.AppendChild(oleSdt);

        // Add a paragraph inside the content control where the OLE object will be placed.
        Paragraph para = new Paragraph(doc);
        oleSdt.AppendChild(para);

        // Use DocumentBuilder to insert the OLE object as an icon.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(para);

        // InsertOleObject expects streams for the OLE data and optional icon.
        using (FileStream oleStream = File.OpenRead(oleSourcePath))
        {
            // No custom icon is provided (null), Aspose.Words will use the default icon.
            Shape oleShape = builder.InsertOleObject(oleStream, "Sample OLE", true, null);

            // Adjust the size of the OLE icon.
            oleShape.Width = 100;
            oleShape.Height = 100;
        }

        // Save the document as DOCX.
        string docxPath = Path.Combine(outputDir, "OleObject.docx");
        doc.Save(docxPath);

        // Convert and save the document as PDF. The OLE object will render as an icon.
        string pdfPath = Path.Combine(outputDir, "OleObject.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
