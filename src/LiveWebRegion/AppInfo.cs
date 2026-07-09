namespace LiveWebRegion
{
    /// <summary>Single source of truth for the version and the update manifest URL.</summary>
    internal static class AppInfo
    {
        public const string Version = "v04";

        // Public raw URL of update.json on the default branch (checked for new versions).
        public const string ManifestUrl =
            "https://raw.githubusercontent.com/michi-burtscher/presentation-web-view/main/update.json";
    }
}
