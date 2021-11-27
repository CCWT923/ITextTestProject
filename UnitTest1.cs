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
    /// IText 7 ������Ŀ
    /// </summary>
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }
        string fileName = @"C:\Users\WuTao\Desktop\HelloWorld.pdf";
        /// <summary>
        /// ����һ���򵥵� PDF �ĵ�
        /// </summary>
        [Test]
        public void HelloWorld()
        {
            using var writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            
            //���һ��ҳ��
            //pdf.AddNewPage(PageSize.A4);
            var document = new Document(pdf);
            Paragraph p0 = new Paragraph("Hello World! ���");
            document.Add(p0);
            TestContext.WriteLine( pdf.GetLastPage().GetPageSize().GetWidth());
            document.Close();
        }

        /// <summary>
        /// ��ҳ��
        /// </summary>
        [Test]
        public void AddPageBreak()
        {
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            pdf.AddNewPage(PageSize.A5);
            //���������һ�������жϣ�������ҳ���СΪ A3
            AreaBreak brk = new AreaBreak(PageSize.A3);
            document.Add(brk);
            //���һ�������жϣ���������Ϊ��һ�����򣬲�ָ��ֽ�ţ�Ĭ��Ϊ A4
            AreaBreak brk2 = new AreaBreak(iText.Layout.Properties.AreaBreakType.NEXT_AREA);
            document.Add(brk2);
            document.Close();
        }

        [Test]
        public void AddParagraph()
        {
            //��������Ķ���
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            Paragraph para = new Paragraph("haha, ");
            para.Add("this is a paragraph in Pdf.");
            //��һ�������������·�����ڶ���������ʾˮƽ������д�� Unicode �ַ���
            PdfFont font1 = PdfFontFactory.CreateFont(@"C:\Users\WuTao\AppData\Local\Microsoft\Windows\Fonts\LXGWWenKai-Light.ttf", PdfEncodings.IDENTITY_H);
            //��һ������ʹ�õ�΢���ź����壬��Ϊ����һ�����壬����ָ��ʹ������ĵ�һ������
            PdfFont font2 = PdfFontFactory.CreateFont(@"C:\windows\fonts\msyh.ttc,0", PdfEncodings.IDENTITY_H);
            Paragraph p2 = new Paragraph("�������ֲ��ԡ������ţ���������������������");
            p2.SetFontSize(25);
            p2.SetFont(font1);
            p2.Add("�������ƺ�֮ˮ���������������������أ������������������׷���������˿ĺ��ѩ��");
            Paragraph p3 = new Paragraph("���������뾡����Īʹ���׿ն��¡������Ҳı����ã�ǧ��ɢ����������");
            //ʹ�� Style ��������
            Style style = new Style();
            style.SetFont(font2);
            style.SetFontSize(30);
            //���ö������ʽ
            p3.AddStyle(style);
            //��������ӵ��ĵ�
            document.Add(para);
            document.Add(p2);
            document.Add(p3);
            document.Close();
        }

        /// <summary>
        /// �б�
        /// </summary>
        [Test]
        public void AddList()
        {
            //��������Ķ���
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);
            //ʵ����һ���б���ָ���б�ֵ��ʽ
            var list = new iText.Layout.Element.List(ListNumberingType.ROMAN_UPPER);
            //Ϊ�б���һ����ʽ
            Style style = new Style();
            //���һ��������
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
        /// �������
        /// </summary>
        [Test]
        public void CreateTable()
        {
            PdfWriter writer = new PdfWriter(fileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);

            //��������п�
            float[] columnWidth = { 200F, 200F, 200F };
            //���� Table ʱ����ָ���ڶ�������Ϊ True���������Խ�������ŵ�ҳ���С������ָ�����п�
            Table table1 = new Table(columnWidth);
            //��������ӵ�Ԫ��
            Cell cell1 = new Cell();
            cell1.Add(new Paragraph("Name"));

            Cell cell2 = new Cell();
            cell2.Add(new Paragraph("Address"));

            //��Ԫ���м���ͼƬ
            var imgData = ImageDataFactory.CreatePng(new System.Uri(@"D:\Documents\My Picture\WallPaper\00 (1).png"));
            Image img = new Image(imgData );
            //��������ͼƬ������߶�
            //img.SetMaxHeight(30);
            img.SetAutoScaleHeight(true); //ͼƬ���ݵ�Ԫ��ߴ��Զ����Ÿ߶ȣ����ȣ�
            Cell cell3 = new Cell();
            cell3.Add(img);

            //�ϲ���Ԫ�����������ֱ�Ϊ�ϲ����С��ϲ�����
            Cell cell4 = new Cell(1, 3);
            cell4.Add(new Paragraph("Merge Cell"));

            //��Ԫ�񱳾�
            Cell cell5 = new Cell();
            cell5.Add(new Paragraph("Background fill"));
            cell5.SetBackgroundColor(ColorConstants.GRAY);
            //���뷽ʽ
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
            //���ó����ӵ�����
            link.SetBold(); //����Ӵ�
            link.SetItalic(); //б��
            link.SetUnderline(); //�»���
            link.SetFontColor(ColorConstants.BLUE); //������ɫ
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
            //��ȡ��ҳ����
            int pagesCount = pdf.GetNumberOfPages();
            for(int i = 1; i <= pagesCount; i++)
            {
                string pageStr = string.Format("Page {0} of {1}", i, pagesCount);
                //��ÿһҳ�����½����һ��ҳ����ʾ
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

            //ˮƽ�ָ��ߣ�����
            LineSeparator ls = new LineSeparator(new DashedLine());
            //ʵ��
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
            //�����
            img.SetMaxWidth(300);
            //����Ϊ���ж���
            img.SetHorizontalAlignment(HorizontalAlignment.CENTER);
            document.Add(img);
            document.Close();
        }
    }
}