using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

class HyphenationPerSection
{
    static void Main()
    {
        // Paths to the source document and hyphenation dictionary files.
        string inputDocPath = @"C:\Docs\MixedLanguages.docx";
        string outputDocPath = @"C:\Docs\MixedLanguages_Hyphenated.docx";

        // Register hyphenation dictionaries for the languages that appear in the document.
        // The dictionaries are in OpenOffice format (*.dic) and must exist at the specified locations.
        Hyphenation.RegisterDictionary("en-US", @"C:\Dictionaries\hyph_en_US.dic");
        Hyphenation.RegisterDictionary("de-CH", @"C:\Dictionaries\hyph_de_CH.dic");
        Hyphenation.RegisterDictionary("fr-FR", @"C:\Dictionaries\hyph_fr_FR.dic");

        // Load the mixed‑language document.
        Document doc = new Document(inputDocPath);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Example logic: assign a language to each section.
        // In a real scenario you would determine the language dynamically,
        // e.g., by inspecting the text or using custom metadata.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            Section section = doc.Sections[i];
            string languageCode;

            // Simple mapping: first section English, second German, third French, repeat if more sections.
            switch (i % 3)
            {
                case 0: languageCode = "en-US"; break;
                case 1: languageCode = "de-CH"; break;
                default: languageCode = "fr-FR"; break;
            }

            // Convert the language code to a Windows LCID.
            int lcid = new CultureInfo(languageCode).LCID;

            // Apply the LCID to every run in the section so that hyphenation uses the correct dictionary.
            foreach (Run run in section.Body.GetChildNodes(NodeType.Run, true))
            {
                run.Font.LocaleId = lcid;
            }
        }

        // Save the document; hyphenation will be applied during the save operation.
        doc.Save(outputDocPath);
    }
}
