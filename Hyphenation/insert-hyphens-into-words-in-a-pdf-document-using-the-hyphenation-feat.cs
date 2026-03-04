using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenatePdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some long text that will wrap onto multiple lines.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
            "Phasellus faucibus velit a ligula fermentum, a iaculis massa aliquet. " +
            "Suspendisse potenti. Curabitur non nulla sit amet nisl tempus convallis quis ac lectus.");

        // Enable automatic hyphenation so that hyphens are inserted at line breaks.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional settings.
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document as PDF; hyphenated words will appear in the output.
        doc.Save("HyphenatedOutput.pdf");
    }
}
