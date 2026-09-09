using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for formatting APIs

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document containing email, phone, and URL.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact us at john.doe@example.com or call 123-456-7890.");
        builder.Writeln("Visit https://www.example.com for more information.");
        builder.Writeln("Alternative email: jane_smith@domain.org, phone: (555) 123 4567, site: http://example.org/page.");

        // Save the document so that we also demonstrate loading from a file.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document from the file system.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Perform sequential replacements: email -> [email protected], phone -> [phone], URL -> [url].
        // -----------------------------------------------------------------
        // Email addresses.
        var emailPattern = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w+\b", RegexOptions.IgnoreCase);
        int emailReplacements = loaded.Range.Replace(emailPattern, "[email protected]", new FindReplaceOptions());
        if (emailReplacements == 0)
            throw new InvalidOperationException("Expected at least one email address replacement.");

        // Phone numbers (simple patterns covering formats like 123-456-7890, (555) 123 4567).
        var phonePattern = new Regex(@"\b(?:\(\d{3}\)\s*|\d{3}[-\s])\d{3}[-\s]\d{4}\b");
        int phoneReplacements = loaded.Range.Replace(phonePattern, "[phone]", new FindReplaceOptions());
        if (phoneReplacements == 0)
            throw new InvalidOperationException("Expected at least one phone number replacement.");

        // URLs (http or https).
        var urlPattern = new Regex(@"\bhttps?://[^\s]+", RegexOptions.IgnoreCase);
        int urlReplacements = loaded.Range.Replace(urlPattern, "[url]", new FindReplaceOptions());
        if (urlReplacements == 0)
            throw new InvalidOperationException("Expected at least one URL replacement.");

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Optional: write a short confirmation to the console (no user interaction required).
        Console.WriteLine($"Replacements completed. Emails: {emailReplacements}, Phones: {phoneReplacements}, URLs: {urlReplacements}.");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
