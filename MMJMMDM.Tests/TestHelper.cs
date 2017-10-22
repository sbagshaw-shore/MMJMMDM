using System.IO;
using System.Reflection;
using System.Resources;

namespace MMJMMDM.Tests
{
    public class TestHelper
    {
        public static string GetContentsOfResourceFile(string resourcePath, Assembly assembly)
        {
            string result;

            using (var stream = assembly.GetManifestResourceStream(resourcePath))
            {
                if (stream == null)
                {
                    throw new MissingManifestResourceException(resourcePath);
                }

                using (var reader = new StreamReader(stream))
                {
                    result = reader.ReadToEnd();
                }
            }

            return result;
        }
    }
}
