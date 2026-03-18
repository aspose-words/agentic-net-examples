using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FormFieldSafeReplace
{
    /// <summary>
    /// Loads a document, checks that the specified range contains no form fields,
    /// and if the check passes performs a find‑and‑replace operation.
    /// </summary>
    /// <param name="inputFile">Path to the source .docx file.</param>
    /// <param name="outputFile">Path where the modified document will be saved.</param>
    /// <param name="searchText">Text to search for.</param>
    /// <param name="replaceText">Text to replace with.</param>
    public static void ReplaceIfNoFormFields(string inputFile, string outputFile, string searchText, string replaceText)
    {
        // Load the document from disk.
        Document doc = new Document(inputFile);

        // Use the whole document range for the validation.
        // The FormFields collection belongs to the Range object.
        bool hasFormFields = doc.Range.FormFields.Count > 0;

        if (hasFormFields)
        {
            // If any form fields are present we abort the replace operation.
            // Depending on requirements you could throw, log, or simply skip.
            Console.WriteLine("The document contains form fields; replacement was skipped.");
            return;
        }

        // No form fields found – safe to perform the replacement.
        // Simple string replace; you could also use FindReplaceOptions if needed.
        int replacementsMade = doc.Range.Replace(searchText, replaceText);
        Console.WriteLine($"Replacements performed: {replacementsMade}");

        // Save the modified document.
        doc.Save(outputFile);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Expecting: inputFile outputFile searchText replaceText
        if (args.Length < 4)
        {
            Console.WriteLine("Usage: <inputFile> <outputFile> <searchText> <replaceText>");
            return;
        }

        string inputFile = args[0];
        string outputFile = args[1];
        string searchText = args[2];
        string replaceText = args[3];

        FormFieldSafeReplace.ReplaceIfNoFormFields(inputFile, outputFile, searchText, replaceText);
    }
}
