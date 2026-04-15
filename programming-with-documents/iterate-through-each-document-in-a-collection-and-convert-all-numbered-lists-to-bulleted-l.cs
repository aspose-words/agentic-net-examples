using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListConversionExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare an output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create sample documents and store them in a collection.
            List<Document> documents = new List<Document>
            {
                CreateNumberedListDocument(),
                CreateMixedListDocument()
            };

            // Process each document: convert numbered lists to bulleted lists.
            int docIndex = 1;
            foreach (Document doc in documents)
            {
                // Ensure a bulleted list template exists in the document (create once).
                List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

                // Iterate over all paragraphs.
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph para in paragraphs)
                {
                    // Check if the paragraph is part of a list.
                    if (para.ListFormat.IsListItem)
                    {
                        // Determine if the current list level uses a numbering style (i.e., not a bullet).
                        ListLevel level = para.ListFormat.ListLevel;
                        if (level != null && level.NumberStyle != NumberStyle.Bullet)
                        {
                            // Switch the paragraph to the bulleted list while preserving its level number.
                            para.ListFormat.List = bulletList;
                            // ListLevelNumber remains the same, so no additional assignment is required.
                        }
                    }
                }

                // Save the modified document.
                string outPath = Path.Combine(outputDir, $"ProcessedDocument{docIndex}.docx");
                doc.Save(outPath);
                docIndex++;
            }
        }

        // Creates a document that contains a simple numbered list.
        private static Document CreateNumberedListDocument()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Sample numbered list:");
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.Writeln("Item 3");
            builder.ListFormat.RemoveNumbers();

            return doc;
        }

        // Creates a document that contains both numbered and bulleted lists.
        private static Document CreateMixedListDocument()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Mixed list example:");
            // Numbered part.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Numbered 1");
            builder.Writeln("Numbered 2");
            builder.ListFormat.RemoveNumbers();

            // Bulleted part.
            builder.ListFormat.ApplyBulletDefault();
            builder.Writeln("Bulleted A");
            builder.Writeln("Bulleted B");
            builder.ListFormat.RemoveNumbers();

            return doc;
        }
    }
}
