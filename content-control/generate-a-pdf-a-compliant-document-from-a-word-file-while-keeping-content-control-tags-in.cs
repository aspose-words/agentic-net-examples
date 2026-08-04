using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string inputDocPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "contentControls.json");
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");

        // -----------------------------------------------------------------
        // 1. Create a sample Word document that contains a plain‑text content control.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some introductory text.
        builder.Writeln("Customer information:");

        // Create an inline plain‑text StructuredDocumentTag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "John Doe"));

        // Insert the content control after the introductory text.
        builder.InsertNode(sdt);
        builder.Writeln(); // Move to the next line.

        // Save the seed document.
        doc.Save(inputDocPath);

        // -----------------------------------------------------------------
        // 2. Load the document, enumerate its content controls and write metadata to JSON.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);

        // Collect information about each StructuredDocumentTag.
        var sdtInfo = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .Cast<StructuredDocumentTag>()
            .Select(tag => new
            {
                Title = tag.Title,
                Tag = tag.Tag,
                Text = tag.GetText().Trim()
            })
            .ToList();

        // Serialize the information to a JSON file.
        string json = JsonConvert.SerializeObject(sdtInfo, Formatting.Indented);
        File.WriteAllText(jsonPath, json);

        // -----------------------------------------------------------------
        // 3. Convert the document to a PDF/A‑1a compliant PDF while preserving the content control.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // PDF/A‑1a compliance (includes PDF/A‑1b plus document structure).
            Compliance = PdfCompliance.PdfA1a,

            // Preserve form fields so that the content control becomes an interactive PDF field.
            PreserveFormFields = true,

            // Use the Tag property of the SDT as the name of the PDF form field.
            UseSdtTagAsFormFieldName = true,

            // Export the document structure (required for PDF/A‑1a, but set explicitly).
            ExportDocumentStructure = true
        };

        // Save the PDF/A document.
        loadedDoc.Save(outputPdfPath, pdfOptions);
    }
}
