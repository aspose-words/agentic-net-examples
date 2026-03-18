using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Define the name of the style to apply to all headings.
        const string targetStyleName = "MyHeadingStyle";

        // Ensure the target style exists; create it if it does not.
        Style targetStyle = doc.Styles[targetStyleName];
        if (targetStyle == null)
        {
            targetStyle = doc.Styles.Add(StyleType.Paragraph, targetStyleName);
            targetStyle.Font.Name = "Arial";
            targetStyle.Font.Size = 14;
            targetStyle.Font.Color = Color.Blue;
            targetStyle.ParagraphFormat.SpaceAfter = 12;
        }

        // Iterate through every paragraph in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Check if the paragraph uses a built‑in heading style (Heading1‑Heading9).
            StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
            if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
            {
                // Apply the custom style.
                para.ParagraphFormat.Style = targetStyle;
            }
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
