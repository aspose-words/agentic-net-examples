using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReporting
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOTX template.
            string sourcePath = @"C:\Docs\Template.dotx";

            // Path where the filtered report will be saved.
            string resultPath = @"C:\Docs\FilteredReport.docx";

            // Keyword to search for inside paragraphs.
            string keyword = "Important";

            // Load the DOTX document (create/load rule).
            Document sourceDoc = new Document(sourcePath);

            // Retrieve all paragraphs from the document (including those in tables, headers, etc.).
            // GetChildNodes returns a live collection; we cast to Paragraph for LINQ usage.
            List<Paragraph> allParagraphs = sourceDoc
                .GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .ToList();

            // Filter paragraphs that contain the specified keyword (LINQ Where + Contains).
            List<Paragraph> matchingParagraphs = allParagraphs
                .Where(p => p.GetText().Contains(keyword, StringComparison.OrdinalIgnoreCase))
                .ToList();

            // Create a new empty document to hold the filtered paragraphs (create rule).
            Document resultDoc = new Document();

            // Remove the default section/paragraph that Aspose.Words adds on creation.
            resultDoc.RemoveAllChildren();

            // Add a new section and body to the result document.
            Section section = new Section(resultDoc);
            resultDoc.AppendChild(section);
            Body body = new Body(resultDoc);
            section.AppendChild(body);

            // Clone each matching paragraph and add it to the result document.
            foreach (Paragraph para in matchingParagraphs)
            {
                // Clone creates a deep copy of the paragraph node.
                Paragraph clonedPara = (Paragraph)para.Clone(true);
                body.AppendChild(clonedPara);
            }

            // Save the filtered report (save rule).
            resultDoc.Save(resultPath);
        }
    }
}
