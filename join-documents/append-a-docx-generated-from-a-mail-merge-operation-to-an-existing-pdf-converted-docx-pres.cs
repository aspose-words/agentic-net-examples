using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create destination DOCX (simulating a PDF‑converted document)
        // -----------------------------------------------------------------
        string destPath = Path.Combine(outputDir, "Destination.docx");
        var destinationDoc = new Document();
        var destBuilder = new DocumentBuilder(destinationDoc);

        // Create a custom style in the destination document
        Style destStyle = destinationDoc.Styles.Add(StyleType.Paragraph, "DestStyle");
        destStyle.Font.Name = "Arial";
        destStyle.Font.Size = 14;
        destStyle.Font.Color = System.Drawing.Color.Blue;

        // Add a paragraph using the custom style
        destBuilder.ParagraphFormat.Style = destStyle;
        destBuilder.Writeln("This is content from the PDF‑converted DOCX (destination).");
        destinationDoc.Save(destPath);

        // -----------------------------------------------------------------
        // 2. Create a mail‑merge template DOCX
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        var templateDoc = new Document();
        var tmplBuilder = new DocumentBuilder(templateDoc);

        // Insert merge fields
        tmplBuilder.Writeln("Dear <<Name>>,");
        tmplBuilder.Writeln("Your appointment is scheduled for <<Date>>.");
        tmplBuilder.Writeln("Thank you.");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Perform mail merge to generate the source DOCX
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        var sourceDoc = new Document(templatePath);
        sourceDoc.MailMerge.Execute(
            new[] { "Name", "Date" },
            new object[] { "John Doe", DateTime.Now.ToString("d") });
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 4. Append the mail‑merged document to the destination document,
        //    preserving destination styles
        // -----------------------------------------------------------------
        destinationDoc.AppendDocument(sourceDoc, ImportFormatMode.UseDestinationStyles);

        // Save the merged result
        string mergedPath = Path.Combine(outputDir, "MergedResult.docx");
        destinationDoc.Save(mergedPath);

        // -----------------------------------------------------------------
        // 5. Validation
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        var mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        // Verify that content from both documents is present
        if (!mergedText.Contains("PDF‑converted DOCX") && !mergedText.Contains("John Doe"))
            throw new InvalidOperationException("Merged document does not contain expected content.");

        // Simple check for at least two sections (original + appended)
        if (mergedDoc.Sections.Count < 2)
            throw new InvalidOperationException("Merged document does not contain the expected number of sections.");

        // Program completed successfully
    }
}
