using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("Inline equation example:");

        // Insert a field of type Equation (EQ) and keep the field code.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = field as FieldEQ;
        if (fieldEQ == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Move the builder to the field separator and write a simple EQ argument.
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(@"\f(1,2)"); // Represents the fraction 1/2.

        // Update the field so that the EQ code is processed.
        field.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = fieldEQ.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        CompositeNode parent = field.Start.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Field start does not have a valid composite parent.");

        parent.InsertBefore(officeMath, field.Start);
        field.Remove();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        // Count the OfficeMath nodes in the document.
        int mathCount = doc.GetChildNodes(NodeType.OfficeMath, true).Count;

        // Report the result.
        Console.WriteLine($"Document saved to '{outputPath}'. OfficeMath nodes count: {mathCount}");
    }
}
