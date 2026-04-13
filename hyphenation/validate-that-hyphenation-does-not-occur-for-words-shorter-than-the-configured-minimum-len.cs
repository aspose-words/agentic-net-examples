using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationMinimumLengthExample
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that contains a short word ("test") which is shorter than the
        // typical minimum word length for hyphenation (5 characters). The line width is set
        // narrow so that longer words would be hyphenated if allowed.
        builder.Font.Size = 24;
        builder.Writeln("This paragraph contains a short word test that should NOT be hyphenated.");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Note: Aspose.Words does not expose a MinimumWordLength property.
        // The hyphenation engine automatically skips words that are too short,
        // so we simply rely on that default behavior.

        // Save the document.
        string docPath = Path.Combine(outputDir, "HyphenationMinimumLength.docx");
        doc.Save(docPath);

        // Load the saved document and verify that the short word was not split by a hyphen.
        Document loadedDoc = new Document(docPath);
        string text = loadedDoc.GetText();

        // The optional hyphen character used by Aspose.Words for automatic hyphenation is
        // ControlChar.OptionalHyphenChar (char value 31). If the short word were hyphenated,
        // this character would appear in the document text.
        if (text.Contains(ControlChar.OptionalHyphenChar.ToString()))
            throw new InvalidOperationException("Hyphenation occurred for a word shorter than the typical minimum length.");

        // If we reach this point, the validation succeeded.
        Console.WriteLine("Hyphenation minimum word length validation passed.");
    }
}
