using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

namespace ListConversionExample
{
    public class Program
    {
        // Creates a sample document that contains a simple numbered list.
        private static void CreateSampleDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a default numbered list.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Numbered item 1");
            builder.Writeln("Numbered item 2");
            builder.Writeln("Numbered item 3");
            builder.ListFormat.RemoveNumbers();

            // Save the document.
            doc.Save(filePath);
        }

        // Converts all numbered list items in the given document to use a bulleted list.
        private static void ConvertNumberedListsToBullets(Document doc)
        {
            // Create a single bulleted list that will be reused for all paragraphs.
            List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

            // Get all paragraphs in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph paragraph in paragraphs)
            {
                // Process only paragraphs that are part of a list.
                if (!paragraph.ListFormat.IsListItem)
                    continue;

                // Determine if the current list item uses a numbered style.
                // Numbered lists have a NumberStyle other than Bullet.
                List currentList = paragraph.ListFormat.List;
                int levelNumber = paragraph.ListFormat.ListLevelNumber;
                NumberStyle style = currentList.ListLevels[levelNumber].NumberStyle;

                if (style != NumberStyle.Bullet)
                {
                    // Switch the paragraph to the bulleted list while preserving its level.
                    paragraph.ListFormat.List = bulletList;
                    paragraph.ListFormat.ListLevelNumber = levelNumber;
                }
            }
        }

        public static void Main()
        {
            // Prepare a collection of document file paths.
            List<string> sourceFiles = new List<string>
            {
                "Document1.docx",
                "Document2.docx"
            };

            // Ensure each source document exists by creating sample files.
            foreach (string file in sourceFiles)
            {
                CreateSampleDocument(file);
            }

            // Process each document: load, convert lists, and save the result.
            foreach (string sourcePath in sourceFiles)
            {
                // Load the document.
                Document doc = new Document(sourcePath);

                // Convert numbered lists to bulleted lists.
                ConvertNumberedListsToBullets(doc);

                // Save the modified document with a new name.
                string outputPath = System.IO.Path.GetFileNameWithoutExtension(sourcePath) + "_Bulleted.docx";
                doc.Save(outputPath);
            }
        }
    }
}
