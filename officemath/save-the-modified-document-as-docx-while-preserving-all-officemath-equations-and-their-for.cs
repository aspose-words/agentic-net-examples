using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;   // Needed for OfficeMath

public class OfficeMathSaveExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some regular text.
        builder.Writeln("This paragraph precedes an equation.");

        // Insert first equation using deterministic EQ-field bootstrap.
        InsertEquation(builder, @"\f(1,2)"); // Simple fraction 1/2

        // Insert another paragraph.
        builder.Writeln("Another paragraph before second equation.");

        // Insert second equation.
        InsertEquation(builder, @"\r(3,x)"); // Simple root expression

        // Modify a paragraph's text to demonstrate document changes while preserving equations.
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        if (firstParagraph.Runs.Count > 0)
        {
            firstParagraph.Runs[0].Text = "Modified first paragraph text.";
        }

        // Save the document as DOCX.
        string outputPath = "ModifiedDocument.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validation: ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output DOCX file was not created.");

        // Validation: ensure OfficeMath nodes are still present.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        if (officeMathNodes.Count == 0)
            throw new InvalidOperationException("No OfficeMath equations were found after saving.");

        // Output counts to console (non‑interactive).
        Console.WriteLine($"Document saved to '{outputPath}'. OfficeMath count: {officeMathNodes.Count}");
    }

    // Helper method to insert an equation via EQ field and convert it to a real OfficeMath node.
    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an empty Equation field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Write the EQ argument after the field separator.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(eqArgument);

        // Update the field so that the EQ argument is processed.
        field.Update();

        // Convert the field to an OfficeMath node.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        Node fieldStart = fieldEq.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

        // Remove the original field (both start and end nodes).
        fieldEq.Remove();

        // Move the builder to the end of the inserted OfficeMath to continue building.
        builder.MoveTo(officeMath);
        builder.Writeln(); // Ensure subsequent content starts on a new line.
    }
}
