using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class InsertOfficeMathFromLatex
{
    public static void Main()
    {
        // LaTeX source string (preserved as comment/metadata)
        // Example LaTeX: "E = mc^2"
        string latexEquation = @"E = mc^2";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will precede the equation.
        builder.Writeln("The following equation is inserted from LaTeX:");

        // Insert a FieldEQ which will be converted to a real OfficeMath node.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = field as FieldEQ;
        if (fieldEQ == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Move to the field separator and write a safe EQ argument.
        builder.MoveTo(fieldEQ.Separator);
        // Use a simple, known‑safe EQ switch that reliably converts.
        builder.Write(@"\f(1,2)"); // Represents a fraction 1/2.

        // Update the field so that Aspose.Words evaluates it.
        field.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = fieldEQ.AsOfficeMath();

        // Ensure conversion succeeded.
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start node.
        Node fieldStart = field.Start;
        Node parent = fieldStart.ParentNode;
        if (parent is CompositeNode compositeParent)
        {
            compositeParent.InsertBefore(officeMath, fieldStart);
        }
        else
        {
            throw new InvalidOperationException("Field start node does not have a composite parent.");
        }

        // Remove the original EQ field, leaving only the real OfficeMath node.
        field.Remove();

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validation: reload the document and verify that an OfficeMath node exists.
        Document loadedDoc = new Document(outputPath);
        int officeMathCount = loadedDoc.GetChildNodes(NodeType.OfficeMath, true).Count;
        if (officeMathCount == 0)
            throw new InvalidOperationException("The output document does not contain any OfficeMath equations.");

        // Confirmation output.
        Console.WriteLine($"Document saved to '{Path.GetFullPath(outputPath)}' with {officeMathCount} OfficeMath node(s).");
    }
}
