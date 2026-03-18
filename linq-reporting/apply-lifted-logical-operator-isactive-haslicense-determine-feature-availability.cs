using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Determine whether the feature is logically active in the application.
        bool isActive = GetFeatureActivationState();

        // Check if Aspose.Words metered license is present and valid.
        bool hasLicense = Metered.IsMeteredLicensed();

        // Lifted logical AND: feature is available only when both conditions are true.
        bool featureAvailable = isActive && hasLicense;

        Console.WriteLine($"Feature available: {featureAvailable}");
    }

    // Placeholder for whatever logic decides if the feature should be active.
    static bool GetFeatureActivationState()
    {
        // Replace with real activation logic.
        return true;
    }
}
