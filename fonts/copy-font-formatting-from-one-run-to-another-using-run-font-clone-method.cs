using System;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Drawing; // For Aspose.Drawing.Color creation

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to work with.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create the source run with distinct font formatting.
        Run sourceRun = new Run(doc, "Source text. ");
        sourceRun.Font.Name = "Courier New";
        sourceRun.Font.Size = 24;
        sourceRun.Font.Bold = true;
        // Convert Aspose.Drawing.Color to System.Drawing.Color as required by Font.Color.
        sourceRun.Font.Color = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Blue.ToArgb());
        paragraph.AppendChild(sourceRun);

        // Create the destination run that will receive the copied formatting.
        Run destinationRun = new Run(doc, "Copied formatting text.");
        paragraph.AppendChild(destinationRun);

        // Copy font properties from the source run to the destination run.
        destinationRun.Font.Name = sourceRun.Font.Name;
        destinationRun.Font.Size = sourceRun.Font.Size;
        destinationRun.Font.Bold = sourceRun.Font.Bold;
        // Font.Color of sourceRun is already a System.Drawing.Color, so copy directly.
        destinationRun.Font.Color = sourceRun.Font.Color;

        // Optional validation of the copied properties.
        Debug.Assert(destinationRun.Font.Name == "Courier New");
        Debug.Assert(destinationRun.Font.Size == 24);
        Debug.Assert(destinationRun.Font.Bold);
        Debug.Assert(destinationRun.Font.Color.ToArgb() == System.Drawing.Color.Blue.ToArgb());

        // Save the document to the local file system.
        doc.Save("CopyFontFormatting.docx");
    }
}
