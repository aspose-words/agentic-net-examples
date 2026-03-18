using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace CombinePageAndHeadingFlags
{
    class Program
    {
        static void Main(string[] args)
        {
            // Adjust these paths to point to your actual input and output folders.
            string dataDir = Path.Combine(Environment.CurrentDirectory, "Data");
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");

            // Ensure the directories exist.
            Directory.CreateDirectory(dataDir);
            Directory.CreateDirectory(outputDir);

            // Load an existing Word document.
            string inputPath = Path.Combine(dataDir, "InputDocument.docx");
            Document doc = new Document(inputPath);

            // Configure HTML save options to split the output into separate parts.
            // Combine the PageBreak and HeadingParagraph flags so that each part starts on a new page.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph;

            // Optionally limit heading levels that trigger a split (default is 2).
            saveOptions.DocumentSplitHeadingLevel = 2;

            // Save the document; each part will be written to its own HTML file.
            string outputPath = Path.Combine(outputDir, "OutputDocument.html");
            doc.Save(outputPath, saveOptions);

            Console.WriteLine($"Document split and saved to: {outputPath}");
        }
    }
}
