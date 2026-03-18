using System;
using Aspose.Words;
using Aspose.Words.Lists;

class ApplyCustomListTemplate
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Create a custom list based on a predefined template (e.g., NumberDefault).
        // This list will have 9 levels by default.
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Example customizations for the first two levels (optional).
        ListLevel level0 = customList.ListLevels[0];
        level0.Font.Color = System.Drawing.Color.DarkBlue;
        level0.Font.Size = 12;
        level0.NumberStyle = NumberStyle.UppercaseLetter;
        level0.NumberFormat = "\x0000.";

        ListLevel level1 = customList.ListLevels[1];
        level1.Font.Color = System.Drawing.Color.DarkGreen;
        level1.Font.Size = 12;
        level1.NumberStyle = NumberStyle.LowercaseRoman;
        level1.NumberFormat = "\x0000)";

        // Iterate over all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Check if the paragraph is part of any list.
            if (para.ListFormat.IsListItem)
            {
                // Apply the custom list to the paragraph.
                para.ListFormat.List = customList;

                // Ensure the list level does not exceed the maximum (8, which is the 9th level).
                if (para.ListFormat.ListLevelNumber > 8)
                    para.ListFormat.ListLevelNumber = 8;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
