using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Data.Xml.Dom;
using Windows.Storage;
using Windows.Storage.Streams;

namespace BibleProcess
{
    public class data
    {
        string localFileName = "BProcessSetting.xml";
        string dataFile = @"ms-appx:///DataModel/coreData_EN.xml";
        string localFileFullPath = string.Empty;

        public data()
        {
            localFileFullPath = string.Format(@"ms-appdata:///local/{0}", localFileName);
            if (App.SystemLanguage == 0)
            {
                dataFile = @"ms-appx:///DataModel/coreData.xml";
            }
            else if(App.SystemLanguage==1)
            {
                dataFile = @"ms-appx:///DataModel/coreData_EN.xml";
            }
        }

        public static int CurrentCHP { get; set; }

        public async Task<int[]> GetTasksIndex()
        {
            int[] result = new int[2];

            Uri fileUrl = new Uri(localFileFullPath);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);

            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(sStream);

            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            if (settings.Count >= 1)
            {
                int chpsPerTime = Int32.Parse(settings[0].SelectSingleNode("setEach").InnerText.Trim());
                int currIndex = Int32.Parse(settings[0].SelectSingleNode("currentIndex").InnerText.Trim());
                int total = Int32.Parse(settings[0].SelectSingleNode("total").InnerText.Trim());
                if (total > 1189)
                    total = 1189;
                int fromIndex = 0;
                int toIndex = 0;
                if (currIndex == total)
                {
                    fromIndex = 1;
                    toIndex = fromIndex + chpsPerTime - 1;
                    currIndex = toIndex;
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

                result[0] = fromIndex;
                result[1] = toIndex;
                CurrentCHP = toIndex;
            }
            return result;
        }

        // no longer used, used GetTasksIndex and GetDisplayNameByIndex
        private async Task<string[]> GetTasksResult()
        {
            int[] _tasksIndex = await GetTasksIndex();
            string from = await GetDisplayNameByIndex(_tasksIndex[0]);
            string to = await GetDisplayNameByIndex(_tasksIndex[1]);

            return new string[] { from, to };
        }

        public async Task<string> GetDisplayNameByIndex(int index)
        {
            string result = string.Empty;
            Uri fileUrl = new Uri(dataFile);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);

            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(sStream);
            if (index < 1)
            {
                return string.Empty;
            }

            XmlNodeList contents = xDoc.GetElementsByTagName("Contents");
            if (contents.Count >= 1)
            {
                var _node = contents[0].SelectSingleNode(string.Format(@"//book[@index='{0}']", index));
                result = _node.SelectSingleNode("displayName").InnerText.Trim();
            }

            return result;
        }

        // nolonger used, user GetMaliSetting
        public async Task<int> GetChpsPerTime()
        {
            Uri fileUrl = new Uri(localFileFullPath);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);

            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(sStream);

            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            int result = 0;
            if (settings.Count >= 1)
            {
                result = Int32.Parse(settings[0].SelectSingleNode("setEach").InnerText.Trim());
            }
            return result;
        }

        // nolonger used, user GetMaliSetting
        public async Task<int> GetallPass()
        {
            Uri fileUrl = new Uri(localFileFullPath);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);

            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(sStream);
            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            int result = 0;
            if (settings.Count >= 1)
            {
                result = Int32.Parse(settings[0].SelectSingleNode("allPass").InnerText.Trim());
            }

            return result;
        }

        // no longer used, used GetTasksIndex, then Index[0]-1 is the current index
        private async Task<int> GetCurrentChpIndex()
        {
            Uri fileUrl = new Uri(localFileFullPath);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);

            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(sStream);
            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            int result = 0;
            if (settings.Count >= 1)
            {
                result = Int32.Parse(settings[0].SelectSingleNode("currentIndex").InnerText.Trim());
            }

            return result;
        }

        // no longer used, replaced by using GetCurrentChpIndex and GetDisplayNameByIndex
        private async Task<string> GetCurrentChp()
        {
            Uri fileUrl = new Uri(localFileFullPath);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);

            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(sStream);
            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            int result = 0;
            if (settings.Count >= 1)
            {
                int currentIndex = Int32.Parse(settings[0].SelectSingleNode("currentIndex").InnerText.Trim());
                result = currentIndex;
            }

            if (result == 0)
                return string.Empty;
            else
            {
                string currChpDisplay = await GetDisplayNameByIndex(result);
                return currChpDisplay;
            }
        }

        // nolonger used, user GetMaliSetting
        private async Task<int> GetFontSize()
        {
            Uri fileUrl = new Uri(localFileFullPath);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);

            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(sStream);

            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            int value = 0;
            if (settings.Count >= 1)
            {
                value = int.Parse(settings[0].SelectSingleNode("fontSize").InnerText.Trim());
            }

            return value;
        }

        /// <summary>
        /// ChpsPerTime, allpass, FontSize, Toast, language
        /// </summary>
        public async Task<int[]> GetMaliSetting()
        {
            Uri fileUrl = new Uri(localFileFullPath);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);

            int[] value = new int[] { 3, 0, 1, 0,0 };
            XmlDocument xDoc = new XmlDocument();

            xDoc.LoadXml(sStream);

            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");

            if (settings.Count >= 1)
            {
                value[0] = int.Parse(settings[0].SelectSingleNode("setEach").InnerText.Trim());
                value[1] = int.Parse(settings[0].SelectSingleNode("allPass").InnerText.Trim());
                value[2] = int.Parse(settings[0].SelectSingleNode("fontSize").InnerText.Trim());
                value[3] = int.Parse(settings[0].SelectSingleNode("toast").InnerText.Trim());
                value[4] = int.Parse(settings[0].SelectSingleNode("language").InnerText.Trim());
            }

            return value;
        }
        /// <summary>
        /// get chapter content
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<string> GetChpsContent(string xmlFile)
        {
            Uri fileUrl = new Uri(xmlFile);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);
            string content = Converter.DoConversion(sStream);
            return content;
        }

        public async Task<bool> SetChpsPerEachAndFontSize(int chpsPerEachValue, int fontSizeValue, int toastValue, int languageValue)
        {
            bool isSaved = false;
            Uri fileUrl = new Uri(localFileFullPath);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);

            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(sStream);

            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            string result = string.Empty;
            if (settings.Count >= 1)
            {
                settings[0].SelectSingleNode("setEach").InnerText = chpsPerEachValue.ToString();
                settings[0].SelectSingleNode("fontSize").InnerText = fontSizeValue.ToString();
                settings[0].SelectSingleNode("toast").InnerText = toastValue.ToString();
                settings[0].SelectSingleNode("language").InnerText = languageValue.ToString();
                try
                {
                    await xDoc.SaveToFileAsync(sFile);
                    isSaved = true;
                }
                catch { }
            }

            return isSaved;
        }

        public async Task<bool> SetCurrentCHP(int value)
        {
            bool complete = false;
            Uri fileUrl = new Uri(localFileFullPath);
            StorageFile sFile = await StorageFile.GetFileFromApplicationUriAsync(fileUrl);
            string sStream = await FileIO.ReadTextAsync(sFile);

            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(sStream);

            XmlNodeList settings = xDoc.GetElementsByTagName("Settings");
            string result = string.Empty;
            if (settings.Count >= 1)
            {
                if (value == 1189)
                {
                    int _allPass = Int32.Parse(settings[0].SelectSingleNode("allPass").InnerText.Trim());
                    _allPass++;
                    settings[0].SelectSingleNode("allPass").InnerText = _allPass.ToString();
                }
                settings[0].SelectSingleNode("currentIndex").InnerText = value.ToString();
                await xDoc.SaveToFileAsync(sFile);
                complete = true;
            }
            return complete;
        }

        public async Task<bool> InitSetting()
        {
            bool isSaved = false;
            string originalSetting = @"ms-appx:///DataModel/Setting.xml";
            Uri originalSettingURL = new Uri(originalSetting);
            StorageFile originalSettingFile = await StorageFile.GetFileFromApplicationUriAsync(originalSettingURL);
            string settingStream = await FileIO.ReadTextAsync(originalSettingFile);
            // load content to XmlDocument
            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(settingStream);
           
            StorageFile localFile = await ApplicationData.Current.LocalFolder.CreateFileAsync(localFileName,CreationCollisionOption.OpenIfExists);

            try
            {
                // replace the Setting.xml in local folder with the original version.
                await xDoc.SaveToFileAsync(localFile);
                isSaved = true;
            }
            catch
            {
                isSaved = false;
            }

            return isSaved;
        }

        public async Task<bool> IsFileExist()
        {
            StorageFile sFile = null;
            bool fileExists = false;

            try
            {
                sFile = await StorageFile.GetFileFromApplicationUriAsync(new Uri(localFileFullPath));
                fileExists = true;
            }
            catch (FileNotFoundException ex)
            {
                fileExists = false;
            }
            return fileExists;
        }

        public async Task<bool> CheckFileCompleteness()
        {
            bool isFileCompleteness = false;
            Uri localSettingURL = new Uri(localFileFullPath);
            StorageFile localFile = await StorageFile.GetFileFromApplicationUriAsync(localSettingURL);

            string sStream = await FileIO.ReadTextAsync(localFile);
            if (sStream.Length > 220)
            {
                isFileCompleteness = true;
            }

            return isFileCompleteness;
        }

    }
}
