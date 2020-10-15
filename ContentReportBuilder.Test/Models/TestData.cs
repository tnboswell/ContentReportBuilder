using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.ElementModels;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Models.Styles;
using System;
using System.Collections.Generic;

namespace ContentReportBuilder.Test.Models
{

    public class TestConstants
    {
        public const string OUTPUT_FILE = "output";
    }

    public enum TestType
    {
        BASIC
    }

    public class TestTableRow
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public string Value { get; set; }
    }

    public class TestData
    {

        public static DocumentModel BarChartBasic(DocumentBuilderType type, DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        { 
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new ChartData
            {
                Title = "Test Bar Chart Basic",
                Series = new List<ChartSeries> { 
                    new ChartSeries( "Bar A", 10), 
                    new ChartSeries("Bar B", 1),
                    new ChartSeries("Bar C", 5),
                }
            };
            var element = new BarChartElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel LineChartBasic(DocumentBuilderType type, DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();

            var lineA = new List<SeriesItem> {
                new SeriesItem(1,1),
                new SeriesItem(2,2),
                new SeriesItem(3,3),
                new SeriesItem(4,4),
                new SeriesItem(5,5),
            };

            var lineB = new List<SeriesItem> {
                new SeriesItem(1,5),
                new SeriesItem(2,4),
                new SeriesItem(3,3),
                new SeriesItem(4,2),
                new SeriesItem(5,1)
            };

            var lineC = new List<SeriesItem> {
                new SeriesItem(1,10),
                new SeriesItem(1,5),
                new SeriesItem(1,1),
                new SeriesItem(1,17),
                new SeriesItem(1,10),
            };

            var data = new ChartData
            {
                Title = "Test Line Chart Basic",
                Series = new List<ChartSeries> {
                    new ChartSeries("Line A", lineA),
                    new ChartSeries("Line B", lineB),
                    new ChartSeries("Line C", lineC),
                }
            };
            var element = new LineChartElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel PieChartBasic(DocumentBuilderType type, DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();

            var data = new ChartData
            {
                Title = "Test Pie Chart Basic",
                Series = new List<ChartSeries> {
                    new ChartSeries("Slice A", 5),
                    new ChartSeries("Slice B", 5),
                    new ChartSeries("Slice C", 5),
                }
            };
            var element = new PieChartElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel HeadingBasic(DocumentBuilderType type, DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new HeadingData {  Text = "Test Heading" };
            var element = new HeadingElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel ListBasic(DocumentBuilderType type, DocumentMetaData docMeta, string sectionTitle, bool ordered, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new ListData { 
                Ordered = ordered, 
                Items = new List<ListItemData> { 
                    new ListItemData{ Text = "Item 1" },
                    new ListItemData{ Text = "Item 2" },
                    new ListItemData{ Text = "Item 3" }
                }
            };
            var element = new ListElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel ParagraphBasic(DocumentBuilderType type, DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new ParagraphData { Text = "Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots" +
                " in a piece of classical Latin literature from 45 BC, making it over 2000 years old. Richard McClintock, a Latin professor" +
                " at Hampden-Sydney College in Virginia, looked up one of the more obscure Latin words, consectetur, from a Lorem Ipsum passage," +
                " and going through the cites of the word in classical literature, discovered the undoubtable source. Lorem Ipsum comes from sections" +
                " 1.10.32 and 1.10.33 of 'de Finibus Bonorum et Malorum' (The Extremes of Good and Evil) by Cicero, written in 45 BC." +
                " This book is a treatise on the theory of ethics, very popular during the Renaissance. The first line of Lorem Ipsum," +
                " 'Lorem ipsum dolor sit amet..', comes from a line in section 1.10.32." };
            var element = new ParagraphElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }


        public static DocumentModel TableBasic(DocumentBuilderType type, DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new TableData { Rows = new List<object> {
                new TestTableRow{ Id = 1, Title = "Cat", Value = "Nala" },
                new TestTableRow{ Id = 2, Title = "Dog", Value = "Mike" },
                new TestTableRow{ Id = 3, Title = "Bird", Value = "Jacob" },
            } };
            var element = new TableElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel ImgUrlBasic( DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new WebImageData("https://encrypted-tbn0.gstatic.com/images?q=tbn%3AANd9GcSASCKoxCFxxl37ULd24tTXtXhPwZcbEE1yU_QpYKuOIEARXWU9&usqp=CAU");
            var element = new ImageElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel ImgLocalBasic(DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new LocalImageData("testImage.jpg");
            var element = new ImageElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel HeaderParagraphBasic(DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new HeaderData { Paragraph = new ParagraphData { Text = "Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots" } };
            //var data = new LocalImageData("testImage.jpg");
            var element = new HeaderElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel HeaderImageBasic(DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new HeaderData { Image = new LocalImageData("testImage.jpg") };
            var element = new HeaderElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel FooterParagraphBasic(DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new FooterData { Paragraph = new ParagraphData { Text = "Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots" } };
            var element = new FooterElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel FooterImageBasic(DocumentMetaData docMeta, string sectionTitle, bool isExcel = false)
        {
            var docModel = TestData.GetDocument(TestType.BASIC, docMeta, isExcel);
            var sectionMetaData = new SectionMetaData(sectionTitle);
            var elementList = new List<DocumentModelElement>();
            var data = new FooterData { Image = new LocalImageData("testImage.jpg") };
            var element = new FooterElementModel(data, null, null);
            elementList.Add(element);
            docModel.AddDocumentSection(new DocumentModelSection(sectionMetaData, elementList));
            return docModel;
        }

        public static DocumentModel GetDocument(TestType type, DocumentMetaData docMeta, bool isExcel = false)
        {
            return new DocumentModel(new DocumentStyle(isExcel), docMeta);
        }

    }
}
