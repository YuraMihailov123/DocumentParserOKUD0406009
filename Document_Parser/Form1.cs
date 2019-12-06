using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using SautinSoft.Document;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Document_Parser
{
    public partial class Form1 : Form
    {
        public XDocument xdoc = new XDocument();

        public string date = "";
        public string adoptionDate = "";        
        public string folderPath = @"C:\Users\sodrk\Desktop\net";

        public Form1()
        {
            InitializeComponent();
            SplitDocumentByPages();
        }

        public void SplitDocumentByPages()
        {
            string filePath = @"C:\Users\sodrk\Desktop\net\doc.rtf";
            DocumentCore dc = DocumentCore.Load(filePath);
            
            DocumentPaginator dp = dc.GetPaginator();

            for (int i = 0; i < dp.Pages.Count; i++)
            {
                DocumentPage page = dp.Pages[i];
                Directory.CreateDirectory(folderPath);

                page.Save(folderPath + @"\Page - " + (i+1).ToString() + ".pdf", SautinSoft.Document.SaveOptions.PdfDefault);
            }
            LoadPages(dp.Pages.Count);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
        }

        public void LoadPages(int pagesCount)
        {
            for(int i = 0; i < pagesCount; i++)
            {
                List<SautinSoft.Document.Tables.TableRow> rowContent = new List<SautinSoft.Document.Tables.TableRow>();
                DocumentCore dc = DocumentCore.Load(folderPath +@"\Page - " + (i+1).ToString() +".pdf");
                foreach (SautinSoft.Document.Tables.TableRow run in dc.GetChildElements(true, ElementType.TableRow))
                {
                    rowContent.Add(run);
                };
                foreach (SautinSoft.Document.Section run in dc.GetChildElements(true, ElementType.Section))
                {
                    date = run.Blocks[run.Blocks.Count - 2].Content.ToString().Replace("\r\n", "").Substring(4,10);
                    break;
                };
                foreach (SautinSoft.Document.Paragraph run in dc.GetChildElements(true, ElementType.Paragraph))
                {
                    if (run.Content.ToString().Contains("Дата принятия уполномоченным банком"))
                    {
                        adoptionDate = run.Inlines[7].Content.ToString();
                        break;
                    }
                };
                ParseAndGetInfo(rowContent,i+1);
            }
        }

        public void ParseAndGetInfo(List<SautinSoft.Document.Tables.TableRow> pageContent,int pageNum)
        {
            bool isStartedReadTAbleInfo = false;
            List<string> toXML = new List<string>();
            for (int i = 0; i < pageContent.Count; i++)
            {
                if (pageContent[i].Cells.Count < 12 && isStartedReadTAbleInfo)
                    break;
                if (isStartedReadTAbleInfo)
                {
                    toXML.Add(pageContent[i].Cells[1].Content.ToString().Replace("\r\n",""));
                    toXML.Add(pageContent[i].Cells[2].Content.ToString().Replace("\r\n", ""));
                    toXML.Add(pageContent[i].Cells[4].Content.ToString().Replace("\r\n", ""));
                    toXML.Add(pageContent[i].Cells[5].Content.ToString().Replace("\r\n", ""));
                    toXML.Add(pageContent[i].Cells[6].Content.ToString().Replace("\r\n", ""));
                    toXML.Add(pageContent[i].Cells[7].Content.ToString().Replace("\r\n", ""));
                }
                if (pageContent[i].Cells.Count == 12)
                {
                    if (pageContent[i].Cells[0].Content.ToString().Replace("\r\n", "") == "1" && pageContent[i].Cells[1].Content.ToString().Replace("\r\n", "") == "2")
                    {
                        isStartedReadTAbleInfo = true;
                    }
                }
                
            }
            AddDocumentRowToXML(toXML,pageNum);
            
        }

        public void AddDocumentRowToXML(List<string> attributes,int pageNum) {
            xdoc = new XDocument();
            XElement document = new XElement("Document");
            XAttribute pathPDF = new XAttribute("PDFPath", folderPath + @"\Page - " + (pageNum).ToString() + ".pdf");
            XAttribute dateAttr = new XAttribute("DT", date);
            XAttribute adoptDateAttr = new XAttribute("AdoptionDate", adoptionDate);
            document.Add(pathPDF);
            document.Add(dateAttr);
            document.Add(adoptDateAttr);
            for (int i = 0; i < attributes.Count; i += 6)
            {
                XElement documentRow = new XElement("DocumentRow");
                XElement notNum = new XElement("NotificationNumber", attributes[i]);
                XElement date = new XElement("DT", attributes[i+1]);
                XElement opCode = new XElement("OperationCode", attributes[i + 2]);
                XElement curCode = new XElement("CurrencyCode", attributes[i + 3]);
                XElement sum = new XElement("Sum", attributes[i + 4]);
                XElement docNum = new XElement("DocumentnNumber", attributes[i + 5]);
                documentRow.Add(notNum);
                documentRow.Add(date);
                documentRow.Add(opCode);
                documentRow.Add(curCode);
                documentRow.Add(sum);
                documentRow.Add(docNum);

                document.Add(documentRow);
            }
            xdoc.Add(document);
            xdoc.Save(folderPath + @"\Page - " + (pageNum).ToString() + ".xml");
            
        }

    }
}
