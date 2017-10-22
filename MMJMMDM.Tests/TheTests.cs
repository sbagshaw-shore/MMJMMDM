using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MMJMMDM.Tests
{
    [TestClass]
    public class TheTests
    {
        private static HackIt _hackit;

        [ClassInitialize]
        public static void Setup(TestContext testContext)
        {
            _hackit = new HackIt();
        }

        [TestMethod]
        public void FileWithMultipleParticipantsReturnsCorrectNumberOfStrings()
        {
            var result = _hackit.GetStringsFromFiles(new List<FileInfo>()
            {
                new FileInfo("C:\\Dev\\MMJMMDM\\MMJMMDM.Tests\\Resources\\TestStripped.xlsx")
            });

            Assert.AreEqual(44, result.Count());
        }

        [TestMethod]
        public void EmptyListOfFilesReturnsEmptyListOfStrings()
        {
            var result = _hackit.GetStringsFromFiles(new List<FileInfo>());
            
            Assert.AreEqual(0, result.Count());
        }

        [TestMethod]
        public void TwoFilesReturnsTwoStrings()
        {
            var result = _hackit.GetStringsFromFiles(new List<FileInfo>()
            {
                new FileInfo("C:\\Dev\\MMJMMDM\\MMJMMDM.Tests\\Resources\\Numerical Download_MITCH.xlsx"),
                new FileInfo("C:\\Dev\\MMJMMDM\\MMJMMDM.Tests\\Resources\\Numerical Download_MITCH2.xlsx")
            });

            Assert.AreEqual(2, result.Count());
        }

        [TestMethod]
        public void FileHasCorrectNumberOfFieldsReturned1()
        {
            var result = _hackit.GetStringsFromFiles(new List<FileInfo>()
            {
                new FileInfo("C:\\Dev\\MMJMMDM\\MMJMMDM.Tests\\Resources\\Numerical Download_MITCH.xlsx")
            }).First();

            var splitResult = result.Split(',');

            Assert.AreEqual(125, splitResult.Count());
        }

        [TestMethod]
        public void FileHasCorrectNumberOfFieldsReturned2()
        {
            var result = _hackit.GetStringsFromFiles(new List<FileInfo>()
            {
                new FileInfo("C:\\Dev\\MMJMMDM\\MMJMMDM.Tests\\Resources\\Numerical Download_MITCH2.xlsx")
            }).First();

            var splitResult = result.Split(',');

            Assert.AreEqual(125, splitResult.Count());
        }
    }
}
