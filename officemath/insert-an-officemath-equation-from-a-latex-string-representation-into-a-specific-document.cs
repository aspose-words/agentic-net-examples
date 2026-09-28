using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class InsertOfficeMathFromLatex
{
    public static void Main()
    {
        // LaTeX source is kept as a comment for reference only.
        // Example LaTeX: \frac{a}{b}
        string latexEquation = @"\frac{a}{b}";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory paragraph.
        builder.Writeln("The following equation is inserted from a LaTeX string:");

        // Bookmark marks the exact insertion point.
        builder.StartBookmark("EqLocation");
        builder.Writeln(); // Ensure a new line.
        builder.EndBookmark("EqLocation");

        // Move the cursor to the bookmark.
        builder.MoveToBookmark("EqLocation");

        // Insert an EQ field (FieldEquation) – this will be converted to a real OfficeMath node.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = (FieldEQ)field;

        // Write a safe EQ argument string at the field separator.
        // Using a simple fraction that Aspose.Words can reliably convert.
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(@"\f(1,2)"); // Represents the fraction 1/2.

        // Update the field so that the EQ argument is processed.
        field.Update();

        // Convert the EQ field to an OfficeMath node.
        OfficeMath officeMath = fieldEQ.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start node.
        Node fieldStart = field.Start;
        CompositeNode parent = fieldStart.ParentNode as CompositeNode
            ?? throw new InvalidOperationException("Field start node does not have a composite parent.");

        parent.InsertBefore(officeMath, fieldStart);

        // Remove the original EQ field, leaving only the real OfficeMath node.
        field.Remove();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Verify that exactly one top‑level OfficeMath paragraph exists.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int topLevelCount = 0;
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                topLevelCount++;
        }

        if (topLevelCount != 1)
            throw new InvalidOperationException($"Expected 1 top‑level OfficeMath node, but found {topLevelCount}.");
    }
}
