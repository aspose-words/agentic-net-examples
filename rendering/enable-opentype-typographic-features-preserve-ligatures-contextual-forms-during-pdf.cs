using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace OpenTypePreserveExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();

            // Add a paragraph with sample text that demonstrates OpenType features (e.g., ligatures).
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Name = "Times New Roman";
            builder.Writeln("fi ﬂ ﬁ f i"); // contains ligature characters

            // Preserve OpenType font formatting features during PDF conversion.
            doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = false;

            // Save the document as PDF.
            doc.Save("Output.pdf", SaveFormat.Pdf);
        }
    }
}
