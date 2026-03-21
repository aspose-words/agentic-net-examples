using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitExample
{
    class Program
    {
        static void Main()
        {
            // Create a simple document with three sections programmatically.
            Document doc = new Document();
            doc.RemoveAllChildren(); // Ensure the document starts empty.

            for (int i = 1; i <= 3; i++)
            {
                // Create a new section with a body.
                Section section = new Section(doc);
                Body body = new Body(doc);
                section.AppendChild(body);
                doc.Sections.Add(section);

                // Add a paragraph with some text to the section.
                Paragraph para = new Paragraph(doc);
                para.AppendChild(new Run(doc, $"This is content of section {i}."));
                body.AppendChild(para);
            }

            // Configure save options to split the document at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
            };

            // Use a temporary folder for the output to avoid hard‑coded paths.
            string outputFolder = Path.Combine(Path.GetTempPath(), "AsposeSplitOutput");
            Directory.CreateDirectory(outputFolder);

            // Base file name for the first part; Aspose.Words will generate additional part files.
            string baseFileName = Path.Combine(outputFolder, "SplitPart.html");

            // Save the document. Because DocumentSplitCriteria is set, the document will be split
            // into multiple HTML files (one per section) in the same folder.
            doc.Save(baseFileName, saveOptions);

            Console.WriteLine($"Document split completed. Files are located in: {outputFolder}");
        }
    }
}
