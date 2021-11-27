using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Action;
using iText.Kernel.Pdf.Canvas.Draw;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using NUnit.Framework;

namespace ITextTestProject
{
    /// <summary>
    /// IText 7 测试项目
    /// </summary>
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }
        string fileName = @"C:\Users\WuTao\Desktop\HelloWorld.pdf";
        /// <summary>
        /// 创建一个简单的 PDF 文档
        /// </summary>
        [Test]
        public void HelloWorld()
        {
            using var writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            
            //添加一个页面
            //pdf.AddNewPage(PageSize.A4);
            var document = new Document(pdf);
            Paragraph p0 = new Paragraph("Hello World! 你好");
            document.Add(p0);
            TestContext.WriteLine( pdf.GetLastPage().GetPageSize().GetWidth());
            document.Close();
        }

        /// <summary>
        /// 分页符
        /// </summary>
        [Test]
        public void AddPageBreak()
        {
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            pdf.AddNewPage(PageSize.A5);
            //创建并添加一个区域中断，并设置页面大小为 A3
            AreaBreak brk = new AreaBreak(PageSize.A3);
            document.Add(brk);
            //添加一个区域中断，设置类型为下一个区域，不指定纸张，默认为 A4
            AreaBreak brk2 = new AreaBreak(iText.Layout.Properties.AreaBreakType.NEXT_AREA);
            document.Add(brk2);
            document.Close();
        }

        [Test]
        public void AddParagraph()
        {
            //创建必须的对象
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            Paragraph para = new Paragraph("haha, ");
            para.Add("this is a paragraph in Pdf.");
            //第一个参数是字体的路径，第二个参数表示水平方向上写的 Unicode 字符。
            PdfFont font1 = PdfFontFactory.CreateFont(@"C:\Users\WuTao\AppData\Local\Microsoft\Windows\Fonts\LXGWWenKai-Light.ttf", PdfEncodings.IDENTITY_H);
            //第一个参数使用的微软雅黑字体，因为这是一套字体，所以指定使用里面的第一个字体
            PdfFont font2 = PdfFontFactory.CreateFont(@"C:\windows\fonts\msyh.ttc,0", PdfEncodings.IDENTITY_H);
            Paragraph p2 = new Paragraph("中文文字测试。标点符号：；，。、！？：《》·");
            p2.SetFontSize(25);
            p2.SetFont(font1);
            p2.Add("君不见黄河之水天上来，奔流到海不复回；君不见高堂明镜悲白发，朝如青丝暮成雪。");
            Paragraph p3 = new Paragraph("人生得意须尽欢，莫使金樽空对月。天生我材必有用，千金散尽还复来。");
            //使用 Style 设置字体
            Style style = new Style();
            style.SetFont(font2);
            style.SetFontSize(30);
            //设置段落的样式
            p3.AddStyle(style);
            //将段落添加到文档
            document.Add(para);
            document.Add(p2);
            document.Add(p3);
            document.Close();
        }

        /// <summary>
        /// 列表
        /// </summary>
        [Test]
        public void AddList()
        {
            //创建必须的对象
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            //实例化一个列表，并指定列表值样式
            var list = new iText.Layout.Element.List(ListNumberingType.ROMAN_UPPER);
            //为列表定义一个样式
            Style style = new Style();
            //添加一个左缩进
            style.SetMarginLeft(20);
            list.AddStyle(style);
            list.Add("C++");
            list.Add("C#");
            list.Add("Java");
            list.Add("Python");
            Paragraph p = new Paragraph("Programming Language: ");
            document.Add(p);
            document.Add(list);
            document.Close();
        }

        /// <summary>
        /// 创建表格
        /// </summary>
        [Test]
        public void CreateTable()
        {
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);

            //定义表格的列宽
            float[] columnWidth = { 200F, 200F, 200F };
            //创建 Table 时可以指定第二个参数为 True，这样可以将表格缩放到页面大小，忽略指定的列宽。
            Table table1 = new Table(columnWidth);
            //向表格中添加单元格
            Cell cell1 = new Cell();
            cell1.Add(new Paragraph("Name"));

            Cell cell2 = new Cell();
            cell2.Add(new Paragraph("Address"));

            //单元格中加入图片
            var imgData = ImageDataFactory.CreatePng(new System.Uri(@"D:\Documents\My Picture\WallPaper\00 (1).png"));
            Image img = new Image(imgData );
            //可以设置图片的最大宽高度
            //img.SetMaxHeight(30);
            img.SetAutoScaleHeight(true); //图片根据单元格尺寸自动缩放高度（或宽度）
            Cell cell3 = new Cell();
            cell3.Add(img);

            //合并单元格，两个参数分别为合并的行、合并的列
            Cell cell4 = new Cell(1, 3);
            cell4.Add(new Paragraph("Merge Cell"));

            //单元格背景
            Cell cell5 = new Cell();
            cell5.Add(new Paragraph("Background fill"));
            cell5.SetBackgroundColor(ColorConstants.GRAY);
            //对齐方式
            cell5.SetTextAlignment(TextAlignment.CENTER);
            cell5.SetVerticalAlignment(VerticalAlignment.MIDDLE);
            cell5.SetHeight(50);

            Cell cell6 = new Cell(1, 1)
                .Add(new Paragraph("Cell"))
                .SetTextAlignment(TextAlignment.RIGHT);

            table1.AddCell(cell1);
            table1.AddCell(cell2);
            table1.AddCell(cell3);
            table1.AddCell(cell4);
            table1.AddCell(cell5);
            table1.AddCell(cell6);

            document.Add(table1);

            document.Close();
        }

        [Test]
        public void CreateHyperlink()
        {
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);

            Link link = new Link(" Click Here ", PdfAction.CreateURI("https://www.baidu.com"));
            Paragraph pHyperlink = new Paragraph("Please");
            //设置超链接的属性
            link.SetBold(); //字体加粗
            link.SetItalic(); //斜体
            link.SetUnderline(); //下划线
            link.SetFontColor(ColorConstants.BLUE); //字体颜色
            pHyperlink.Add(link);
            pHyperlink.Add(" open baidu.");

            document.Add(pHyperlink);
            document.Close();
        }

        [Test]
        public void AddPageNumber()
        {
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            pdf.AddNewPage();
            pdf.AddNewPage();
            //获取总页面数
            int pagesCount = pdf.GetNumberOfPages();
            for(int i = 1; i <= pagesCount; i++)
            {
                string pageStr = string.Format("Page {0} of {1}", i, pagesCount);
                //在每一页的右下角添加一个页码显示
                document.ShowTextAligned(
                    new Paragraph(pageStr), 
                    550, 20, i, TextAlignment.CENTER, 
                    VerticalAlignment.TOP, 0);
            }

            document.Close();
        }

        [Test]
        public void AddHorizontalLine()
        {
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);

            //水平分隔线：虚线
            LineSeparator ls = new LineSeparator(new DashedLine());
            //实线
            LineSeparator ls1 = new LineSeparator(new SolidLine());

            document.Add(new Paragraph("P1").SetTextAlignment(TextAlignment.CENTER));
            document.Add(ls);
            document.Add(new Paragraph("p2").SetTextAlignment(TextAlignment.RIGHT));
            document.Add(ls1);

            document.Close();
        }

        [Test]
        public void AddImage()
        {
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);

            string imagePath = @"D:\Documents\My Picture\WallPaper\24D.jpg";
            Image img = new Image(ImageDataFactory.Create(imagePath));
            //最大宽度
            img.SetMaxWidth(300);
            //设置为居中对齐
            img.SetHorizontalAlignment(HorizontalAlignment.CENTER);
            document.Add(img);
            document.Close();
        }
    }
}