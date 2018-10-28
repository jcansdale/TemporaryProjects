using System.Diagnostics;

namespace TemporaryProjects
{
    internal static class Utilities
    {
        public static int VsVersion => Process.GetCurrentProcess().MainModule.FileVersionInfo.FileMajorPart;
    }
}
