using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to host inline content controls.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph ?? new Paragraph(doc);
        if (doc.FirstSection.Body.FirstParagraph == null)
        {
            doc.FirstSection.Body.AppendChild(firstParagraph);
        }

        // ----- Plain text content control: Product Name -----
        StructuredDocumentTag productNameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "ProductName",
            Tag = "product-name"
        };
        productNameSdt.RemoveAllChildren();
        productNameSdt.AppendChild(new Run(doc, "SuperWidget"));
        firstParagraph.AppendChild(productNameSdt);

        // Add a space between controls.
        firstParagraph.AppendChild(new Run(doc, " "));

        // ----- Plain text content control: Product ID -----
        StructuredDocumentTag productIdSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "ProductId",
            Tag = "product-id"
        };
        productIdSdt.RemoveAllChildren();
        productIdSdt.AppendChild(new Run(doc, "SW-001"));
        firstParagraph.AppendChild(productIdSdt);

        // Add a space.
        firstParagraph.AppendChild(new Run(doc, " "));

        // ----- Drop‑down list content control: Category -----
        StructuredDocumentTag categorySdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Category",
            Tag = "category"
        };
        categorySdt.ListItems.Add(new SdtListItem("Electronics", "Electronics"));
        categorySdt.ListItems.Add(new SdtListItem("Tools", "Tools"));
        categorySdt.ListItems.Add(new SdtListItem("Home", "Home"));
        // Set default selected value.
        categorySdt.RemoveAllChildren();
        categorySdt.AppendChild(new Run(doc, "Electronics"));
        firstParagraph.AppendChild(categorySdt);

        // Add a space.
        firstParagraph.AppendChild(new Run(doc, " "));

        // ----- Checkbox content control: In Stock -----
        StructuredDocumentTag inStockSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "InStock",
            Tag = "in-stock",
            Checked = true
        };
        firstParagraph.AppendChild(inStockSdt);

        // Save the document with content controls.
        const string docPath = "product.docx";
        doc.Save(docPath);

        // ----- Extraction for indexing -----
        List<IndexItem> indexItems = new List<IndexItem>();

        foreach (StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                                                .OfType<StructuredDocumentTag>())
        {
            string value;

            // Determine value based on the type of the content control.
            if (sdt.SdtType == SdtType.Checkbox)
            {
                value = sdt.Checked ? "true" : "false";
            }
            else
            {
                // For plain text, rich text, dropdown, etc., use the displayed text.
                value = sdt.GetText().Trim();
            }

            indexItems.Add(new IndexItem
            {
                Title = sdt.Title ?? string.Empty,
                Tag = sdt.Tag ?? string.Empty,
                Value = value
            });
        }

        // Serialize the extracted metadata to JSON.
        string json = JsonConvert.SerializeObject(indexItems, Formatting.Indented);
        const string jsonPath = "index.json";
        File.WriteAllText(jsonPath, json);

        // Output paths for verification (optional).
        Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
        Console.WriteLine($"Index data saved to: {Path.GetFullPath(jsonPath)}");
    }

    private class IndexItem
    {
        public string Title { get; set; } = string.Empty;
        public string Tag { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }
}
