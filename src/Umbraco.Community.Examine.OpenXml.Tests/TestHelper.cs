namespace Umbraco.Community.Examine.OpenXml.Tests;

public static class TestHelper
{
    public static string GetTestFilePath(string fileName)
    {
        return Path.Combine(AppContext.BaseDirectory, "TestFiles", fileName);
    }

    public static Stream GetTestFileStream(string fileName)
    {
        var path = GetTestFilePath(fileName);
        return File.OpenRead(path);
    }
}
