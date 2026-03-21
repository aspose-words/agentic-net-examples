using System;
using System.IO;
using Aspose.Words;

class CIHyphenationUpdater
{
    // Callback that registers hyphenation dictionaries from a folder when the layout engine requests them.
    private class FolderHyphenationCallback : IHyphenationCallback
    {
        private readonly string _folder;

        public FolderHyphenationCallback(string folder)
        {
            _folder = folder;
        }

        public void RequestDictionary(string language)
        {
            // If the dictionary is already registered, nothing to do.
            if (Hyphenation.IsDictionaryRegistered(language))
                return;

            // Expected file name pattern: hyph_en_US.dic for language "en-US".
            string fileName = Path.Combine(_folder, $"hyph_{language.Replace("-", "_")}.dic");

            if (File.Exists(fileName))
            {
                // Register the dictionary using the file‑based overload.
                Hyphenation.RegisterDictionary(language, fileName);
            }
            else
            {
                // Register a null dictionary to suppress further callbacks for this language.
                Hyphenation.RegisterDictionary(language, (string)null);
            }
        }
    }

    static void Main(string[] args)
    {
        // Directory where the CI pipeline places the latest hyphenation dictionaries.
        string dictionariesPath = args.Length > 0 ? args[0] : "./HyphenationDictionaries";

        // Ensure the directory exists to avoid DirectoryNotFoundException.
        Directory.CreateDirectory(dictionariesPath);

        // Pre‑register any dictionaries that already exist in the folder.
        foreach (string file in Directory.GetFiles(dictionariesPath, "hyph_*.dic"))
        {
            // Extract language code from file name, e.g. hyph_en_US.dic -> en-US.
            string nameWithoutExt = Path.GetFileNameWithoutExtension(file); // hyph_en_US
            string langPart = nameWithoutExt.Substring("hyph_".Length);   // en_US
            string language = langPart.Replace('_', '-');                // en-US

            Hyphenation.RegisterDictionary(language, file);
        }

        // Set the callback for on‑demand dictionary loading.
        Hyphenation.Callback = new FolderHyphenationCallback(dictionariesPath);

        // Load a document that will be processed in the CI pipeline.
        Document doc;
        const string templatePath = "Template.docx";
        if (File.Exists(templatePath))
        {
            doc = new Document(templatePath);
        }
        else
        {
            // Create a simple document if the template is missing.
            doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample paragraph used to demonstrate automatic hyphenation in the CI pipeline.");
        }

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the result (PDF is a fixed‑page format that shows hyphenation).
        doc.Save("Processed.pdf");
    }
}
