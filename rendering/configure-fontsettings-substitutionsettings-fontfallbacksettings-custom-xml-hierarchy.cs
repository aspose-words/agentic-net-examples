using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fonts;

class FontFallbackExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize FontSettings and assign them to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Access the FontFallbackSettings object.
        FontFallbackSettings fallbackSettings = fontSettings.FallbackSettings;

        // Define a custom XML hierarchy for fallback fonts.
        // This example maps three Unicode ranges to specific fallback fonts.
        string customFallbackXml = @"<?xml version='1.0' encoding='utf-8'?>
<FallbackSettings>
    <Range Start='0' End='255' FallbackFonts='AllegroOpen' />
    <Range Start='256' End='591' FallbackFonts='AllegroOpen' />
    <Range Start='592' End='1327' FallbackFonts='M+ 2m' />
</FallbackSettings>";

        // Load the custom fallback settings from a memory stream.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(customFallbackXml)))
        {
            fallbackSettings.Load(stream);
        }

        // (Optional) Save the loaded fallback settings to an XML file for verification.
        fallbackSettings.Save("CustomFontFallbackSettings.xml");

        // Save the document to demonstrate that the custom fallback settings are applied.
        doc.Save("DocumentWithCustomFallback.pdf");
    }
}
