using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Sample plain‑text content that contains numbered and bulleted items.
        const string plainText = 
            "Shopping List:\n" +
            "1. Apples\n" +
            "2. Bananas\n" +
            "3. Oranges\n" +
            "\n" +
            "Tasks:\n" +
            "- Clean house\n" +
            "- Pay bills\n";

        // Load the plain‑text into a Document using TxtLoadOptions.
        // AutoNumberingDetection (default true) will try to recognise list items,
        // but we will also demonstrate manual conversion for robustness.
        var loadOptions = new TxtLoadOptions();
        Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(plainText)), loadOptions);

        // Traverse all paragraphs and apply list formatting based on simple patterns.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            string text = para.GetText().Trim();

            // Detect numbered list items like "1. Item".
            if (Regex.IsMatch(text, @"^\d+\.\s+"))
            {
                // Apply a default numbered list to this paragraph.
                para.ListFormat.ApplyNumberDefault();
                // Ensure the list level is the first level (0).
                para.ListFormat.ListLevelNumber = 0;
            }
            // Detect simple bullet items that start with a hyphen.
            else if (Regex.IsMatch(text, @"^-+\s+"))
            {
                para.ListFormat.ApplyBulletDefault();
                para.ListFormat.ListLevelNumber = 0;
            }
        }

        // Save the resulting document with proper Word list structures.
        const string outputPath = "ConvertedList.docx";
        doc.Save(outputPath);
        // The program ends automatically; no user interaction is required.
    }
}
