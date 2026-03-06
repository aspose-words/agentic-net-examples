// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Load a document that contains OpenType features.
        Document doc = new Document("OpenTypeText.docx");

        // Enable advanced typography by assigning a HarfBuzz text shaper factory.
        // This allows the layout engine to use OpenType shaping when rendering.
        doc.LayoutOptions.TextShaperFactory = HarfBuzzTextShaperFactory.Instance;

        // Export the shaped document to PDF (shaping is applied during PDF/XPS export).
        doc.Save("ShapedOutput.pdf");
    }
}
