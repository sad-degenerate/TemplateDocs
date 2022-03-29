using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace TemplateDocs.Test
{
    [TestClass]
    public class ReplacementInDocTest
    {
        [TestMethod]
        public void ReplacmentInDocTestMethod()
        {
            var path = $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\newFile.docx";

            File.Create(path);

            // TODO: Дописать

            Assert.Fail();
        }
    }
}