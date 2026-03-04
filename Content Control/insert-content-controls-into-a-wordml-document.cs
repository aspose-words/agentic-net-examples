using System;
using Aspose.Words;
using Aspose.Words.Markup;

class InsertContentControls
{
    static void Main()
    {
        // Load an existing WORDML (XML) document.
        Document doc = new Document("InputDocument.xml");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (structured document tag).
        // SdtType.PlainText creates a simple text control.
        builder.InsertStructuredDocumentTag(SdtType.PlainText);

        // Write the text that will be inside the content control.
        builder.Write("This text is inside a plain‑text content control.");

        // Optionally close the paragraph after the control.
        builder.Writeln();

        // Insert a rich‑text content control.
        builder.InsertStructuredDocumentTag(SdtType.RichText);
        builder.Write("Rich‑text content control with ");
        builder.InsertField("MERGEFIELD  Author ");
        builder.Writeln(" inside.");

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}
