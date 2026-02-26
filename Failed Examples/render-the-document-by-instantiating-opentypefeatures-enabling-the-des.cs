// ALL ATTEMPTS FAILED. Below is the last generated code.

using Aspose.Words;
using Aspose.Words.Fonts;          // OpenTypeFeatures, OpenTypeFeature
using Aspose.Words.Layout;        // HarfBuzzTextShaperFactory

public class OpenTypeRendering
{
    public static void Run()
    {
        // Load a document that contains text which can benefit from OpenType features.
        Document doc = new Document("OpenType text shaping.docx");

        // Create an OpenTypeFeatures object and enable the desired feature flags.
        OpenTypeFeatures otf = new OpenTypeFeatures();
        // Example: enable Stylistic Set 1 (replace with any required feature).
        otf.EnableFeature(OpenTypeFeature.StylisticSet1);

        // Assign the OpenTypeFeatures to the document's FontSettings.
        doc.FontSettings.OpenTypeFeatures = otf;

        // Use HarfBuzz text shaper so that OpenType features are applied during layout.
        doc.LayoutOptions.TextShaperFactory = HarfBuzzTextShaperFactory.Instance;

        // Render the document to PDF (OpenType features will be applied).
        doc.Save("Document.OpenTypeRendered.pdf");
    }
}
