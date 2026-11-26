using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathTypeDemo
{
    public static void Main()
    {
        // Create a sample document containing a fraction and a radical equation.
        string sourcePath = "Sample.docx";
        CreateSampleDocument(sourcePath);

        // Load the document.
        Document doc = new Document(sourcePath);

        // Retrieve all OfficeMath nodes.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Report the MathObjectType of each node.
        foreach (OfficeMath om in officeMathNodes)
        {
            string typeDescription = om.MathObjectType switch
            {
                MathObjectType.Fraction => "Fraction",
                MathObjectType.Radical => "Radical",
                _ => "Other"
            };

            Console.WriteLine($"OfficeMath node: MathObjectType = {om.MathObjectType}, Determined type = {typeDescription}");
        }

        // Save the processed document.
        string outputPath = "Processed.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The output file '{outputPath}' was not created.");
    }

    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph describing the fraction.
        builder.Writeln("Fraction equation:");
        InsertEquation(builder, @"\f(1,2)");

        // Insert a paragraph describing the radical.
        builder.Writeln("Radical equation:");
        InsertEquation(builder, @"\r(3,x)");

        // Save the sample document.
        doc.Save(filePath, SaveFormat.Docx);

        // Validate that the source file was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"The source file '{filePath}' was not created.");
    }

    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        if (field is not FieldEQ eqField)
            throw new InvalidOperationException("Failed to create a FieldEQ.");

        // Move to the separator and write the EQ argument.
        builder.MoveTo(eqField.Separator);
        builder.Write(eqArgument);

        // Update the field so that it can be converted to OfficeMath.
        field.Update();

        // Convert the field to a real OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();

        // Insert the OfficeMath node before the field start and remove the field.
        if (officeMath != null)
        {
            Node startNode = eqField.Start;
            if (startNode?.ParentNode is CompositeNode parent)
            {
                parent.InsertBefore(officeMath, startNode);
                eqField.Remove();
            }
            else
            {
                throw new InvalidOperationException("Unable to locate a valid parent for the OfficeMath node.");
            }
        }
        else
        {
            throw new InvalidOperationException("EQ field conversion to OfficeMath returned null.");
        }
    }
}
