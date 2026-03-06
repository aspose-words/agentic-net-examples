// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Layout;

class OpenTypeRendering
{
    static void Main()
    {
        // Load the source document (replace the path with your actual file)
        Document doc = new Document("input.docx");

        // Create an OpenTypeFeatures object and enable the required OpenType feature flags
        OpenTypeFeatures otf = new OpenTypeFeatures();
        otf.EnableFeature(OpenTypeFeatureTag.StandardLigatures);          // example feature
        otf.EnableFeature(OpenTypeFeatureTag.ContextualAlternates);      // example feature

        // Assign the OpenTypeFeatures to the document's FontSettings
        doc.FontSettings.OpenTypeFeatures = otf;

        // Use HarfBuzz text shaper so that OpenType features are applied during layout/rendering
        doc.LayoutOptions.TextShaperFactory = HarfBuzzTextShaperFactory.Instance;

        // Render the document as PDF
        doc.Save("output.pdf");
    }
}
