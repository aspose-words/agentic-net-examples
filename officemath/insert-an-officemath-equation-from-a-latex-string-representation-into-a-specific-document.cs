using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Tables; // For CompositeNode base class

public class Program
{
    public static void Main()
    {
        // LaTeX source string (preserved as metadata)
        // Note: Aspose.Words does not parse LaTeX directly; we keep it as a comment.
        string latex = @"\frac{a}{b}";

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an introductory paragraph.
        builder.Writeln("Paragraph before equation.");

        // Mark the insertion point with a bookmark.
        builder.StartBookmark("InsertHere");
        builder.Writeln("Insert equation here:");
        builder.EndBookmark("InsertHere");

        // Move the builder to the bookmark start.
        builder.MoveToBookmark("InsertHere");

        // Insert an EQ field (deterministic bootstrap for OfficeMath).
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Write a safe EQ argument (simple fraction) into the field separator.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(@"\f(1,2)"); // Represents a fraction; used as a placeholder for the LaTeX equation.

        // Update the field so that Aspose.Words evaluates the EQ argument.
        field.Update();

        // Convert the EQ field to a real OfficeMath node.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        Node fieldStart = fieldEq.Start;
        Node parent = fieldStart.ParentNode;
        if (parent is CompositeNode compositeParent)
        {
            compositeParent.InsertBefore(officeMath, fieldStart);
        }
        else
        {
            throw new InvalidOperationException("Field start does not have a composite parent.");
        }

        // Remove the original EQ field.
        fieldEq.Remove();

        // Add a concluding paragraph.
        builder.Writeln();
        builder.Writeln("Paragraph after equation.");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Reload the document and verify the OfficeMath node exists.
        Document loaded = new Document(outputPath);
        NodeCollection mathNodes = loaded.GetChildNodes(NodeType.OfficeMath, true);
        if (mathNodes.Count == 0)
            throw new InvalidOperationException("No OfficeMath nodes were found in the saved document.");

        // Optionally set display type for top-level equations.
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                om.DisplayType = OfficeMathDisplayType.Display;
        }

        // Save any formatting changes.
        loaded.Save(outputPath, SaveFormat.Docx);
    }
}
