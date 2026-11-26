using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a temporary directory for sample documents and output.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create sample documents with numbered lists.
        string[] sourceFiles =
        {
            Path.Combine(workDir, "Doc1.docx"),
            Path.Combine(workDir, "Doc2.docx")
        };

        for (int i = 0; i < sourceFiles.Length; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a numbered list with three items.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln($"Item {i + 1}.A");
            builder.Writeln($"Item {i + 1}.B");
            builder.Writeln($"Item {i + 1}.C");
            builder.ListFormat.RemoveNumbers();

            doc.Save(sourceFiles[i]);
        }

        // Process each document: convert numbered lists to bulleted lists.
        foreach (string filePath in sourceFiles)
        {
            Document doc = new Document(filePath);

            // Create a bulleted list template once per document.
            List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.ListFormat.IsListItem &&
                    para.ListFormat.List != null &&
                    para.ListFormat.List.ListLevels[0].NumberStyle != NumberStyle.Bullet)
                {
                    // Assign the bulleted list to the paragraph.
                    para.ListFormat.List = bulletList;
                }
            }

            // Save the converted document.
            string outputPath = Path.Combine(
                workDir,
                Path.GetFileNameWithoutExtension(filePath) + "_Bulleted.docx");

            doc.Save(outputPath);
        }

        // The program ends here; output files are saved in the Work folder.
    }
}
