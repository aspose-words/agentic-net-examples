using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "LinkControl",
            Tag = "link-control"
        };

        // Add a paragraph that will hold the hyperlink.
        Paragraph para = new Paragraph(doc);
        sdt.AppendChild(para);

        // Move the builder to the newly created paragraph.
        builder.MoveTo(para);

        // Insert a hyperlink field inside the content control.
        builder.Font.Color = System.Drawing.Color.Blue;
        builder.Font.Underline = Underline.Single;
        FieldHyperlink hyperlink = (FieldHyperlink)builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);
        builder.Font.ClearFormatting();

        // Append the content control to the document body.
        doc.FirstSection.Body.AppendChild(sdt);

        // Save the document as DOCX.
        const string docxPath = "LinkControl.docx";
        doc.Save(docxPath);

        // Convert the document to PDF.
        const string pdfPath = "LinkControl.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF back and verify the hyperlink target URL.
        Document pdfDoc = new Document(pdfPath);
        FieldHyperlink loadedHyperlink = null;

        foreach (Field field in pdfDoc.Range.Fields)
        {
            if (field.Type == FieldType.FieldHyperlink)
            {
                loadedHyperlink = (FieldHyperlink)field;
                break;
            }
        }

        if (loadedHyperlink != null)
        {
            Console.WriteLine("Hyperlink target URL: " + loadedHyperlink.Address);
        }
        else
        {
            Console.WriteLine("No hyperlink found in the converted document.");
        }
    }
}
