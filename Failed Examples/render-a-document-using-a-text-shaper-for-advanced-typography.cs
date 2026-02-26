// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Shaping;

class AdvancedTypographyExample
{
    static void Main()
    {
        // Load a document that contains OpenType features (e.g., ligatures, contextual forms).
        Document doc = new Document("OpenTypeText.docx");

        // Assign a text shaper factory that uses HarfBuzz for advanced typography rendering.
        // This enables OpenType shaping when the document is laid out (e.g., during PDF export).
        doc.LayoutOptions.TextShaperFactory = HarfBuzzTextShaperFactory.Instance;

        // Export the document to PDF – the layout will now apply OpenType shaping.
        doc.Save("OpenTypeShaped.pdf");
    }
}
