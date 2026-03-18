using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Folder for generated files.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the two versions of the document.
        string noHyphenationPath = Path.Combine(artifactsDir, "NoHyphenation.docx");
        string hyphenationPath   = Path.Combine(artifactsDir, "Hyphenation.docx");
        string diffPath          = Path.Combine(artifactsDir, "LayoutDifferences.docx");

        // -----------------------------------------------------------------
        // 1. Create a base document with a long paragraph that will wrap.
        // -----------------------------------------------------------------
        Document baseDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(baseDoc);
        builder.Font.Size = 24;

        // Sample text long enough to cause line breaks.
        string sampleText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris " +
                            "nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in " +
                            "reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla " +
                            "pariatur. Excepteur sint occaecat cupidatat non proident, sunt in " +
                            "culpa qui officia deserunt mollit anim id est laborum.";

        builder.Writeln(sampleText);
        // Save the document with hyphenation disabled (default).
        baseDoc.Save(noHyphenationPath);

        // -----------------------------------------------------------------
        // 2. Load the same document and enable automatic hyphenation.
        // -----------------------------------------------------------------
        Document hyphenatedDoc = new Document(noHyphenationPath);
        hyphenatedDoc.HyphenationOptions.AutoHyphenation = true;          // Turn on hyphenation.
        hyphenatedDoc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch zone (optional).
        hyphenatedDoc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Limit consecutive hyphens.
        hyphenatedDoc.Save(hyphenationPath);

        // -----------------------------------------------------------------
        // 3. Compare the two documents to highlight layout differences.
        // -----------------------------------------------------------------
        Document original = new Document(noHyphenationPath);
        Document revised   = new Document(hyphenationPath);

        // The Compare method adds revision marks to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Save the comparison result which shows where hyphenation changed line breaks.
        original.Save(diffPath);
    }
}
