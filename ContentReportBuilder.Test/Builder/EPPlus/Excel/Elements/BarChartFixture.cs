using ContentReportBuilder.Models;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Test.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    [TestClass]
    public class BarChartFixture
    {
        [TestInitialize]
        public void OnTestInitialize()
        {
            _SystemUnderTest = null;
        }

        private ExcelBarChart _SystemUnderTest;
        public ExcelBarChart SystemUnderTest
        {
            get
            {
                if (_SystemUnderTest == null)
                {
                    _SystemUnderTest = new ExcelBarChart();
                }

                return _SystemUnderTest;
            }
        }

        [TestMethod]
        public void BarChart_File_Basic()
        {
            // arrange
            var expected = true;
            var docType = DocumentBuilderType.Excel;
            IDocumentDirector director = new ContentReportBuilder.EPPlus.DocumentDirector();
            var docMeta = new DocumentMetaData("BarChartBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Bar Chart Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.BarChartBasic(docType, docMeta, sectionTitle), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void BarChart_Stream_Basic()
        {
            // arrange
            var docType = DocumentBuilderType.Excel;
            IDocumentDirector director = new ContentReportBuilder.EPPlus.DocumentDirector();
            var docMeta = new DocumentMetaData("BarChartBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Bar Chart Basic Section";
            var docModel = TestData.BarChartBasic(docType, docMeta, sectionTitle);
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
