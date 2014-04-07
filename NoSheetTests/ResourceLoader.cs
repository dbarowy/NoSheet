using System;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace NoSheetTests
{
    public static class ResourceLoader
    {
        /// <summary>
        /// This method returns the correct resource path regardless of
        /// output directory, as long as resources are marked "Copy to
        /// Output Directory."
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static string GetResourcePath(string filename)
        {
            string OUTPUT_PATH = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string DATA_DIRECTORY_NAME = "TestData";
            string path = Path.Combine(OUTPUT_PATH, DATA_DIRECTORY_NAME, filename);
            Assert.IsTrue(File.Exists(path), String.Format("Cannot find test data file \"{0}\" at path \"{1}\".", filename, path));
            return path;
        }
    }
}
