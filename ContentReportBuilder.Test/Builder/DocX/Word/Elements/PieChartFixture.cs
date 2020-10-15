using ContentReportBuilder.Models;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Test.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{

    [TestClass]
    public class PieChartFixture
    {

        [TestInitialize]
        public void OnTestInitialize()
        {
            _SystemUnderTest = null;
        }

        private WordPieChart _SystemUnderTest;
        public WordPieChart SystemUnderTest
        {
            get
            {
                if (_SystemUnderTest == null)
                {
                    _SystemUnderTest = new WordPieChart();
                }

                return _SystemUnderTest;
            }
        }

        [TestMethod]
        public void PieChart_File_Basic()
        {
            // arrange
            var expected = true;
            var docType = DocumentBuilderType.Word;
            IDocumentDirector director = new ContentReportBuilder.docX.DocumentDirector();
            var docMeta = new DocumentMetaData("PieChartBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Pie Chart Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.PieChartBasic(docType, docMeta, sectionTitle), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void PieChart_Stream_Basic()
        {
            // arrange
            var docType = DocumentBuilderType.Word;
            IDocumentDirector director = new ContentReportBuilder.docX.DocumentDirector();
            var docMeta = new DocumentMetaData("PieChartBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Pie Chart Basic Section";
            var docModel = TestData.PieChartBasic(docType, docMeta, sectionTitle);
            string errorMsg;
            // act
            var mem = director.BuildStream(docType, docModel, out errorMsg);

            // assert
            Assert.AreNotEqual(mem, null);
            Assert.AreEqual("", errorMsg);
            StreamToFile.WriteStream(docType, docModel, mem);
        }
    }
}
