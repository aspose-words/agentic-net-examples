using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a numeric expression field (e.g., 12345).
        Field field = builder.InsertField("= 12345");

        // Calculate the field result.
        field.Update();

        // Locate the first Run that contains the field result.
        // The field result appears after the field separator node.
        Node node = field.Start;
        while (node != null && !(node is FieldSeparator))
            node = node.NextSibling;

        // Move past the separator to the first result Run.
        if (node is FieldSeparator)
            node = node.NextSibling;

        Run resultRun = node as Run;
        if (resultRun != null)
        {
            // If the result has more than one character, split it so that only the first character gets custom styling.
            string text = resultRun.Text;
            if (text.Length > 1)
            {
                // Create a run for the first character and apply the desired font formatting.
                Run firstCharRun = new Run(doc, text.Substring(0, 1));
                firstCharRun.Font.Color = Color.Red;   // Example: red color.
                firstCharRun.Font.Bold = true;         // Example: bold.

                // Create a run for the remaining characters (no special formatting).
                Run remainingRun = new Run(doc, text.Substring(1));

                // Insert the new runs in place of the original run.
                CompositeNode parent = resultRun.ParentNode as CompositeNode;
                if (parent != null)
                {
                    parent.InsertBefore(firstCharRun, resultRun);
                    parent.InsertAfter(remainingRun, firstCharRun);
                }

                // Remove the original unstyled run.
                resultRun.Remove();
            }
            else
            {
                // Single‑character result – apply styling directly.
                resultRun.Font.Color = Color.Red;
                resultRun.Font.Bold = true;
            }
        }

        // Save the document with the custom‑styled field result.
        doc.Save("CustomFormattedFieldResult.docx");
    }
}
