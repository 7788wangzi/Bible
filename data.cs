using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Data;
using System.Data.OleDb;

namespace Bible
{
    public class data
    {
#if DEBUG
        const string xml = @"..\..\coreData.xml";
#else
        const string xml =@"coreData.xml";
#endif
        const string excel = @"D:\Projects\Windows Phone\temp\Bible\Bible\Book1.xlsx";

        public string connectionString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=No;IMEX=1'", excel);
        public DataSet resultSet = new DataSet();
        public void GetData()
        {
            string strCommandText = string.Format(@"select * from [{0}$]", "Sheet1");
            OleDbDataAdapter oda = new OleDbDataAdapter(strCommandText, connectionString);
            try
            {
                oda.Fill(resultSet);
            }
            catch
            {
                resultSet = null;
            }
            finally
            {
                oda.Dispose();
            }            
        }

        public void WriteXML()
        {
            int curIndex = 0;
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(xml);
            XmlElement root = xDoc.DocumentElement;
            XmlNodeList contents = root.GetElementsByTagName("Contents");
            if (contents.Count >= 1)
            {
                DataTable dt = resultSet.Tables[0];
                int totalRows = dt.Rows.Count;
                for (int i = 0; i < totalRows-1;i++ )
                {
                    string abbr = dt.Rows[i][1].ToString();
                    string name = dt.Rows[i][2].ToString();
                    string type = dt.Rows[i][3].ToString();
                    
                    int num = Int32.Parse(dt.Rows[i][4].ToString());
                    for (int j = 1; j <= num; j++)
                    {
                        curIndex++;
                        string displayname = string.Format(@"{0}第{1}章",name,j);
                        XmlElement newElement = xDoc.CreateElement("book");
                        XmlAttribute xattri = xDoc.CreateAttribute("index");
                        xattri.Value = curIndex.ToString();
                        newElement.Attributes.Append(xattri);
                        //XmlElement index = xDoc.CreateElement("index");
                        //index.InnerText = curIndex.ToString();
                        XmlElement display = xDoc.CreateElement("displayName");
                        display.InnerText = displayname;
                        XmlElement book = xDoc.CreateElement("bookName");
                        book.InnerText = name;
                        XmlElement chp = xDoc.CreateElement("chapter");
                        chp.InnerText = j.ToString();
                        XmlElement abbrE = xDoc.CreateElement("abbreviation");
                        abbrE.InnerText = abbr;
                        XmlElement typeE = xDoc.CreateElement("type");                        
                        typeE.InnerText = type;
                        //newElement.AppendChild(index);
                        newElement.AppendChild(display);
                        newElement.AppendChild(book);
                        newElement.AppendChild(chp);
                        newElement.AppendChild(abbrE);
                        newElement.AppendChild(typeE);
                        contents[0].AppendChild(newElement);
                    }
                }

                xDoc.Save(xml);
            }
        }

        public string GetResult()
        {
            string result = string.Empty;
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(xml);
            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            bool finished = false;
            if (settings.Count >= 1)
            {
                int chpsPerTime = Int32.Parse(settings[0].ChildNodes[0].InnerText);
                int currIndex = Int32.Parse(settings[0].ChildNodes[1].InnerText);
                int total = Int32.Parse(settings[0].ChildNodes[2].InnerText);
                if (total > 1189)
                    total = 1189;
                int fromIndex =0;
                int toIndex =0;
                if (currIndex == total)
                {
                    fromIndex = 1;
                    toIndex = fromIndex + chpsPerTime - 1;
                    currIndex = toIndex;

                    // finished for 1 time
                    finished = true;
                }
                else
                {
                    fromIndex = currIndex + 1;
                    toIndex = fromIndex + chpsPerTime - 1;
                    if (toIndex <= total)
                    {
                        currIndex = toIndex;
                    }
                    else
                    {
                        toIndex = total;
                        currIndex = toIndex;        
                    }
                }

                //save and get
                XmlNodeList contents = xDoc.GetElementsByTagName("Contents");
                if (contents.Count >= 1)
                {
                    XmlNode fromNode = contents[0].SelectSingleNode(string.Format(@"//book[@index='{0}']", fromIndex));
                    string f1 = fromNode.ChildNodes[0].InnerText;
                    XmlNode toNode = contents[0].SelectSingleNode(string.Format(@"//book[@index='{0}']", toIndex));
                    string t1 = toNode.ChildNodes[0].InnerText;
                    result = string.Format(@"from - {0}: to - {1}",f1,t1);
                }
                settings[0].ChildNodes[1].InnerText = currIndex.ToString();
                if (finished)
                {
                    int soFar = Int32.Parse(settings[0].ChildNodes[3].InnerText);
                    soFar++;
                    settings[0].ChildNodes[3].InnerText = soFar.ToString();
                }
                xDoc.Save(xml);
            }
            return result;
        }

        public string GetChpsPerTime()
        {
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(xml);
            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            string result = string.Empty;
            if (settings.Count >= 1)
            {
                int chpsPerTime = Int32.Parse(settings[0].ChildNodes[0].InnerText);
                result = chpsPerTime.ToString();
            }
            return result;
        }

        public string GetallPass()
        {
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(xml);
            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            string result = string.Empty;
            if (settings.Count >= 1)
            {
                int chpsPerTime = Int32.Parse(settings[0].ChildNodes[3].InnerText);
                result = chpsPerTime.ToString();
            }
            return result;
        }

        public void SetChpsPerTime(int value)
        {
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(xml);
            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            string result = string.Empty;
            if (settings.Count >= 1)
            {
                settings[0].ChildNodes[0].InnerText = value.ToString();
                xDoc.Save(xml);
            }
        }

    }
}
