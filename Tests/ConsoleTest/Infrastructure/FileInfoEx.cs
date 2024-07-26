namespace ConsoleTest.Infrastructure;

internal static class FileInfoEx
{
    public static Stream OpenReadWrite(this FileInfo file) => new FileStream(file.FullName, FileMode.Open, FileAccess.Read | FileAccess.Write, FileShare.Read);
}
