using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(),
            "AsposeDemo_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);

        // Create a simple PDF file to work with.
        string pdfPath = Path.Combine(workDir, "sample.pdf");
        Document tempDoc = new Document();

        Paragraph paragraph = new Paragraph(tempDoc);
        Run run = new Run(tempDoc, "Hello, Aspose.Words!");
        paragraph.AppendChild(run);
        tempDoc.FirstSection.Body.AppendChild(paragraph);

        tempDoc.Save(pdfPath, SaveFormat.Pdf);

        // Path where the resulting Markdown file will be saved.
        string markdownPath = Path.Combine(workDir, "sample.md");

        // Create a temporary folder for extracted images.
        string imagesTempFolder = Path.Combine(workDir, "images");
        Directory.CreateDirectory(imagesTempFolder);

        // Load the PDF document.
        Document doc = new Document(pdfPath);

        // Configure Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ImagesFolder = imagesTempFolder,
            ImagesFolderAlias = ".",
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as Markdown.
        doc.Save(markdownPath, saveOptions);

        Console.WriteLine($"Markdown file saved to: {markdownPath}");
    }
}
