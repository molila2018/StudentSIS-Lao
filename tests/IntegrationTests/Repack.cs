// One-off helper: rebuild a .docx from an unpacked directory tree.
// Use: dotnet run -- repack <unpacked-dir> <output-docx>
// Compress-Archive (PowerShell) writes a parent folder entry that breaks docx;
// ZipFile.CreateFromDirectory produces a flat zip that Office accepts.
using System;
using System.IO;
using System.IO.Compression;

namespace StudentSIS.IntegrationTests;

internal static class Repack
{
    public static void Run(string srcDir, string outPath)
    {
        if (!Directory.Exists(srcDir)) { Console.WriteLine($"NOT FOUND: {srcDir}"); return; }
        if (File.Exists(outPath)) File.Delete(outPath);
        // includeBaseDirectory: false → entries are added relative to srcDir
        ZipFile.CreateFromDirectory(srcDir, outPath, CompressionLevel.Optimal, includeBaseDirectory: false);
        Console.WriteLine($"WROTE: {outPath}  ({new FileInfo(outPath).Length} bytes)");
    }
}
