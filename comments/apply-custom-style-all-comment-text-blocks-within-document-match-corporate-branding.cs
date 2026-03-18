using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class ApplyCommentStyle
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Ensure the custom style exists (character style used for comment text).
        const string styleName = "CorporateComment";
        Style commentStyle = doc.Styles[styleName];
        if (commentStyle == null)
        {
            // Create a new character style.
            commentStyle = doc.Styles.Add(StyleType.Character, styleName);
            // Define corporate branding for the comment text.
            commentStyle.Font.Name = "Calibri";
            commentStyle.Font.Size = 10;
            commentStyle.Font.Color = System.Drawing.Color.DarkBlue;
            commentStyle.Font.Bold = true;
        }

        // Iterate over all comment nodes in the document.
        NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment comment in comments)
        {
            // Apply the custom style to every Run inside the comment.
            foreach (Run run in comment.GetChildNodes(NodeType.Run, true))
            {
                run.Font.StyleName = styleName;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
