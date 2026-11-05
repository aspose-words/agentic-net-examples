using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two heading sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample Document");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of chapter 1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of chapter 2.");

        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Configure HtmlSaveOptions to split by heading.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 1,
            DocumentPartSavingCallback = new CustomPartSavingCallback()
        };

        // Save the document; it will be split into separate HTML files.
        string mainFilePath = Path.Combine(outputFolder, "MainDocument.html");
        doc.Save(mainFilePath, saveOptions);

        // After saving, rename the generated split files to the desired convention.
        // Aspose.Words creates files named "MainDocument.html", "MainDocument_1.html", etc.
        var generatedFiles = Directory.GetFiles(outputFolder, "MainDocument*.html")
                                      .Where(f => !f.Equals(mainFilePath, StringComparison.OrdinalIgnoreCase))
                                      .OrderBy(f => f)
                                      .ToList();

        if (generatedFiles.Count == 0)
            throw new InvalidOperationException("No split HTML files were generated.");

        int chapterIndex = 1;
        foreach (string filePath in generatedFiles)
        {
            string newFileName = $"Chapter_{chapterIndex}.html";
            string newPath = Path.Combine(outputFolder, newFileName);
            File.Move(filePath, newPath);
            Console.WriteLine($"Generated: {newFileName}");
            chapterIndex++;
        }
    }

    // Empty callback – required by the API but does not modify naming directly.
    private class CustomPartSavingCallback : IDocumentPartSavingCallback
    {
        public void DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // No custom handling needed; naming will be adjusted after the save operation.
        }
    }
}
