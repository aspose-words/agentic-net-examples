using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;

class ConvertNumberedToBulleted
{
    static void Main()
    {
        // Collection of source document file paths.
        List<string> sourceFiles = new List<string>
        {
            @"C:\Docs\Doc1.docx",
            @"C:\Docs\Doc2.docx",
            // add more paths as needed
        };

        // Destination folder for the converted documents.
        string outputFolder = @"C:\Docs\Converted\";

        foreach (string filePath in sourceFiles)
        {
            // Load the document (lifecycle rule: load).
            Document doc = new Document(filePath);

            // Create a single bulleted list that will replace all numbered lists
            // (lifecycle rule: create).
            List bulletedList = doc.Lists.Add(ListTemplate.BulletDefault);

            // Iterate through every paragraph in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                // Check if the paragraph is part of a list.
                if (!para.ListFormat.IsListItem)
                    continue;

                // Determine the current list level number.
                int levelNumber = para.ListFormat.ListLevelNumber;

                // Retrieve the list level definition from the paragraph's original list.
                ListLevel originalLevel = para.ListFormat.List.ListLevels[levelNumber];

                // If the original list uses a bullet style, skip it (already bulleted).
                if (originalLevel.NumberStyle == NumberStyle.Bullet)
                    continue;

                // Apply the bulleted list to the paragraph, preserving its level.
                para.ListFormat.List = bulletedList;
                para.ListFormat.ListLevelNumber = levelNumber;
            }

            // Save the modified document (lifecycle rule: save).
            string fileName = System.IO.Path.GetFileName(filePath);
            string outputPath = System.IO.Path.Combine(outputFolder, fileName);
            doc.Save(outputPath);
        }
    }
}
