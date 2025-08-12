// Example using MSTest (recommended for .NET projects in Visual Studio)
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpecRebuilder;
using System.IO;

namespace SpecRebuilder.Tests
{
    [TestClass]
    public class SpecRebuilderTest
    {
        [TestMethod]
        public void Rebuild_CreatesOutputFile()
        {
            // Arrange
            string jsonPath = "C:\\Users\\mnewman\\Documents\\Admin\\Spec Templates\\SpecRebuilder_v0.1\\output\\SECTION 00 00 00_working_converted.json";
            string templatePath = "C:\\Users\\mnewman\\Documents\\Admin\\Spec Templates\\SpecRebuilder_v0.1\\examples\\SECTION 00 00 00.docx";
            string outputPath = "C:\\Users\\mnewman\\Documents\\Admin\\Spec Templates\\SpecRebuilder_v0.1\\output\\test_output.docx";

            if (!File.Exists(jsonPath) || !File.Exists(templatePath))
            {
                Assert.Inconclusive("Test files not found.");
            }

            // Act
            var rebuilder = new SpecRebuilder();
            rebuilder.Rebuild(jsonPath, templatePath, outputPath);

            // Assert
            Assert.IsTrue(File.Exists(outputPath), "Output file was not created.");
        }
    }
}
