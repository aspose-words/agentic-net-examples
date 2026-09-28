using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several checkbox content controls.
        Document doc = new Document();

        // First paragraph (default exists) – add a title.
        Paragraph titlePara = doc.FirstSection.Body.FirstParagraph;
        titlePara.AppendChild(new Run(doc, "Sample document with checkbox content controls"));
        titlePara.AppendChild(new Run(doc, Environment.NewLine));

        // Helper to add a checkbox content control in its own paragraph.
        void AddCheckbox(string title, string tag, bool isChecked)
        {
            // Create a new paragraph.
            Paragraph para = new Paragraph(doc);

            // Create the checkbox content control.
            StructuredDocumentTag checkbox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            checkbox.Title = title;
            checkbox.Tag = tag;
            checkbox.Checked = isChecked;

            // Append the checkbox to the paragraph.
            para.AppendChild(checkbox);

            // Append a descriptive run after the checkbox.
            para.AppendChild(new Run(doc, " " + title));

            // Add the paragraph to the document body.
            doc.FirstSection.Body.AppendChild(para);
        }

        // Add sample checkboxes.
        AddCheckbox("Accept Terms", "acceptTerms", true);
        AddCheckbox("Subscribe Newsletter", "subscribeNewsletter", false);
        AddCheckbox("Enable Notifications", "enableNotifications", true);

        // Save the sample document.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document for processing.
        Document loadedDoc = new Document(docPath);

        // Find all checkbox content controls.
        List<StructuredDocumentTag> checkboxControls = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Checkbox)
            .ToList();

        // Prepare CSV lines.
        List<string> csvLines = new List<string>
        {
            "Title,Tag,Checked"
        };

        foreach (StructuredDocumentTag sdt in checkboxControls)
        {
            // Use empty string if Title or Tag is null.
            string title = sdt.Title ?? string.Empty;
            string tag = sdt.Tag ?? string.Empty;
            string checkedValue = sdt.Checked ? "True" : "False";

            // Escape potential commas by surrounding with double quotes.
            string Escape(string value) => $"\"{value.Replace("\"", "\"\"")}\"";

            csvLines.Add($"{Escape(title)},{Escape(tag)},{checkedValue}");
        }

        // Write CSV to file.
        const string csvPath = "checkboxes.csv";
        File.WriteAllLines(csvPath, csvLines);
    }
}
