using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using NPOI.SS.UserModel;
using Siemens.Engineering;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.HW;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using TIA_Add_In_Helper;

namespace TIA_Add_In_ProDiagTool
{
    /// <summary>
    /// 处理逻辑
    /// </summary>
    public static class Handle
    {
        /// <summary>
        /// 删除文件夹及其内容
        /// </summary>
        /// <param name="targetDir"></param>
        public static void DeleteDirectoryAndContents(string targetDir)
        {
            if (!Directory.Exists(targetDir))
            {
                throw new DirectoryNotFoundException($"目录 {targetDir} 不存在。");
            }

            string[] files = Directory.GetFiles(targetDir);
            string[] dirs  = Directory.GetDirectories(targetDir);

            // 先删除所有文件
            foreach (string file in files)
            {
                File.Delete(file);
            }

            // 然后递归删除所有子目录
            foreach (string dir in dirs)
            {
                DeleteDirectoryAndContents(dir);
            }

            // 最后，删除目录本身
            Directory.Delete(targetDir, false);
        }

        /// <summary>
        /// 导出带诊断的实例DB和ProDiagFB
        /// </summary>
        /// <param name="blockGroup">程序块</param>
        /// <param name="folder">导出文件夹</param>
        private static void ExportProDiagWithInsDB(this PlcBlockGroup blockGroup, string folder)
        {
            foreach (PlcBlock plcBlock in blockGroup.Blocks)
            {
                if (plcBlock.ProgrammingLanguage == ProgrammingLanguage.DB)
                {
                    foreach (EngineeringAttributeInfo info in plcBlock.GetAttributeInfos())
                    {
                        if (info.Name == "Supervisions")
                        {
                            if (plcBlock.GetAttribute("Supervisions") != null)
                            {
                                string filePath = $@"{folder}\{plcBlock.Name}.xml";
                                plcBlock.ExportInfo(filePath);
                            }
                        }
                    }
                }

                if (plcBlock.ProgrammingLanguage == ProgrammingLanguage.ProDiag)
                {
                    string filePath = $@"{folder}\{plcBlock.Name}.csv";
                    plcBlock.ExportInfo(filePath);
                }
            }

            foreach (PlcBlockUserGroup subBlockGroup in blockGroup.Groups)
            {
                ExportProDiagWithInsDB(subBlockGroup, folder);
            }
        }

        /// <summary>
        /// 处理诊断名称，获取接口名称
        /// </summary>
        /// <param name="name">诊断名称</param>
        /// <returns></returns>
        private static string GetSupervisionInterface(string name)
        {
            string       str     = string.Empty;
            const string pattern = @"#(.*?)_O_";
            Match        match   = Regex.Match(name, pattern);
            if (match.Success)
            {
                str = match.Groups[1].Value;
                return str;
            }

            return null;
        }

        /// <summary>
        /// 处理Xml文件获得相关数据
        /// </summary>
        /// <param name="file">Xml文件路径</param>
        /// <param name="proDiag">诊断数据</param>
        /// <returns></returns>
        private static List<XmlInfo> Analyze(this string file, List<ProDiagInfo> proDiag)
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(file);
            var xmlInfos = new List<XmlInfo>();

            //获取所有监控的节点
            XmlNodeList blockInstSupervisionNodes = xmlDocument.GetBlockInstSupervisionNodes();
            foreach (XmlNode blockInstSupervisionNode in blockInstSupervisionNodes)
            {
                string dB_Name   = xmlDocument.GetAttribute("Name"); //DB的名称:P3102
                string dB_Number = xmlDocument.GetAttribute("Number"); //DB的编号:3102
                string stateStruct =
                    blockInstSupervisionNode
                        .GetBlockInstSupervision("StateStruct"); //监控的名称:P3102.#Y_DV_M01_Started_Man_O_1
                string number = blockInstSupervisionNode.GetBlockInstSupervision("Number"); //监控的编号:2755
                string blockTypeSupervisionNumber =
                    blockInstSupervisionNode.GetBlockInstSupervision("BlockTypeSupervisionNumber"); //监控的类型:1
                //获取监控的接口节点
                XmlNode xmlNode = xmlDocument.GetInterfaceMember(GetSupervisionInterface(stateStruct));
                string  offset  = int.Parse(xmlNode.GetAttribute("Offset")).CalOffset(); //监控的偏移量:178.0

                // 添加数据
                XmlInfo xmlInfo = new XmlInfo
                {
                    StateStruct = stateStruct,
                    Number      = number,
                    Offset      = offset,
                    DB_Name     = dB_Name,
                    DB_Number   = dB_Number
                };
                //添加报警类型
                switch (blockTypeSupervisionNumber)
                {
                    case "1":
                        xmlInfo.BlockTypeSupervisionNumber = "Alarms";
                        break;
                    case "2":
                        xmlInfo.BlockTypeSupervisionNumber = "Events";
                        break;
                    case "3":
                        break;
                }

                //添加报警文本
                // 获取对应的所有英文 AlarmText
                var englishAlarmTexts = proDiag
                    .Where(p => p.Identification == number 
                                && p.Language == "en-US" 
                                && p.AlarmText.Contains(dB_Name))
                    .Select(p => p.AlarmText)
                    .ToList();

                string longestAlarmText = string.Empty;

                foreach (var item in englishAlarmTexts.Where(item => item.Length > longestAlarmText.Length))
                {
                    longestAlarmText = item;
                } 
                
                // 最终将最长的字符串赋值给xmlInfo.AlarmTextEn
                xmlInfo.AlarmTextEn = longestAlarmText;

                // 获取对应的所有中文 AlarmText
                var chineseAlarmTexts = proDiag
                    .Where(p => p.Identification == number 
                                && p.Language == "zh-CN" 
                                && p.AlarmText.Contains(dB_Name))
                    .Select(p => p.AlarmText)
                    .ToList();
                
                longestAlarmText = string.Empty;

                foreach (var item in chineseAlarmTexts.Where(item => item.Length > longestAlarmText.Length))
                {
                    longestAlarmText = item;
                } 

                // 最终将最长的字符串赋值给xmlInfo.AlarmTextZh
                xmlInfo.AlarmTextZh = longestAlarmText;

                xmlInfos.Add(xmlInfo);
            }

            //按偏移量顺序排序
            xmlInfos.Sort((x, y) => double.Parse(x.Offset).CompareTo(double.Parse(y.Offset)));

            return xmlInfos;
        }

        /// <summary>
        /// 设置报警文本
        /// </summary>
        /// <param name="language">语言</param>
        /// <param name="workbook"></param>
        /// <param name="xmlInfo"></param>
        /// <param name="increase">递增量</param>
        /// <param name="triggerTag"></param>
        private static void SetAlarmText(this IWorkbook workbook, 
            XmlInfo xmlInfo, string language, int increase,string triggerTag)
        {
            //写入英文报警文本
            ISheet sheet = workbook.GetSheet("DiscreteAlarms");
            IRow   row   = sheet.CreateRow(increase);
            row.SetValue("ID",increase.ToString());
            row.SetValue("Name",xmlInfo.StateStruct);
            switch (language)
            {
                case "en-US":
                    row.SetValue($"Alarm text [{language}], Alarm text", xmlInfo.AlarmTextEn);
                    break;
                case "zh-CN":
                    row.SetValue($"Alarm text [{language}], Alarm text", xmlInfo.AlarmTextZh);
                    break;
            }
            row.SetValue("Class",xmlInfo.BlockTypeSupervisionNumber);
            row.SetValue("Trigger tag",triggerTag);
            string[] parts = xmlInfo.Offset.Split('.');
            if (parts.Length > 1)
            {
                string result = parts[1];
                row.SetValue("Trigger bit",result);
            }
        }

        /// <summary>
        /// 获取所有设备信息
        /// </summary>
        /// <param name="projectBase">项目</param>
        /// <returns>设备名称Name、设备Device类</returns>
        public static List<DeviceInfo> GetDeviceInfos(this ProjectBase projectBase)
        {
            //获取所有设备，并把名称和Device写入到数据
            var devices = new List<DeviceInfo>();
            foreach (Device device in projectBase.AllDevices())
            {
                if (device.DeviceItems[0].GetAttribute("TypeIdentifier").ToString().Contains(":6AV2"))
                {
                    DeviceInfo deviceInfo = new DeviceInfo { Name = device.Name, Device = device };
                    devices.Add(deviceInfo);
                }
            }

            return devices;
        }

        /// <summary>
        /// PLC数据处理
        /// </summary>
        /// <param name="deviceItem">设备项</param>
        /// <param name="exportFolder">导出文件夹路径</param>
        /// <param name="xmlInfos">实例DB的信息</param>
        public static void PlcDataAnalyze(this DeviceItem deviceItem,
            string                                        exportFolder,
            out List<XmlInfo>                             xmlInfos
        )
        {
            //查询PLC目标
            PlcSoftware plcSoftware = deviceItem.GetPlcSoftware();
            //导出xml和csv
            PlcBlockGroup plcBlockGroup = plcSoftware.BlockGroup;
            plcBlockGroup.ExportProDiagWithInsDB(exportFolder);

            //初始化数据
            var proDiagInfos = new List<ProDiagInfo>();
            xmlInfos     = new List<XmlInfo>();

            //轮询文件夹中所有相关文件
            if (Directory.Exists(exportFolder))
            {
                //轮询文件夹中所有的csv文件
                string[] csvFiles = Directory.GetFiles(exportFolder, "*.csv");
                // 筛选出文件名包含 en-US 或 zh-CN 的文件
                var filteredFiles = csvFiles.Where(file =>
                    Path.GetFileName(file).Contains("en-US") ||
                    Path.GetFileName(file).Contains("zh-CN"));
                
                //获取报警文本
                foreach (string csv in filteredFiles)
                {
                    using (StreamReader reader = new StreamReader(csv))
                    {
                        var data = reader.Analyze(Path.GetFileName(csv));
                        proDiagInfos.AddRange(data);
                    }
                }

                //轮询文件夹中所有的xml文件
                string[] xmlFiles = Directory.GetFiles(exportFolder, "*.xml");
                //获取xml数据
                foreach (string xml in xmlFiles)
                {
                    xmlInfos.AddRange(xml.Analyze(proDiagInfos));
                }
            }
        }

        /// <summary>
        /// Hmi数据处理
        /// </summary>
        /// <param name="hmiTarget">Hmi目标</param>
        /// <param name="exportFolder">导出文件夹路径</param>
        /// <param name="xmlInfos">实例DB的信息</param>
        /// <param name="tagTable">Hmi变量表</param>
        /// <param name="hmiAlarms_enUS">触摸屏报警文本xlsx(英文)</param>
        /// <param name="hmiAlarms_zhCN">触摸屏报警文本xlsx(中文)</param>
        public static void HmiDataAnalyze(this HmiTarget hmiTarget,
            string exportFolder,
            List<XmlInfo> xmlInfos,
            out XmlDocument tagTable,
            out IWorkbook hmiAlarms_enUS,
            out IWorkbook hmiAlarms_zhCN
            )
        {
            //导出默认变量表
            TagTable defaultTagTable = hmiTarget.TagFolder.DefaultTagTable;
            string   exportPath      = $@"{exportFolder}\{defaultTagTable.Name}.xml";
            defaultTagTable.ExportInfo(exportPath);
            //获取Connection名称
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(exportPath);
            string connection = xmlDocument.GetLinkList("Connection");

            //定义触摸屏报警变量表
            tagTable = new XmlDocument();
            tagTable = HmiSoftware.TagTable();
            tagTable.SetAttribute("Name", "自动生成报警变量表");
            XmlDocument tag = new XmlDocument();
            tag = HmiSoftware.Tag();

            //定义触摸屏报警文本
            hmiAlarms_enUS = HmiSoftware.HmiAlarmText("en-US");
            hmiAlarms_zhCN = HmiSoftware.HmiAlarmText("zh-CN");
            
            //定义报警文本和变量处理的辅助变量
            int    tagCount = 1; //变量的名称递增
            int    tagID    = 1; //变量的ID递增
            int    textID   = 1; //报警文本ID递增
            string old      = null;
            string tagName  = null;
            
            //处理报警文本和变量
            foreach (XmlInfo xmlInfo in xmlInfos)
            {
                //写入Hmi变量表Xml
                //第一个变量
                if (old != xmlInfo.DB_Name)
                {
                    tag.SetId(tagID); //设置ID
                    //计算和设置地址%DB1.DBW1
                    string logicalAddress = xmlInfo.Offset.GetLogicalAddress(xmlInfo.DB_Number);
                    tag.SetAttribute("LogicalAddress", logicalAddress);
                    //计算和设置名称Pxxxx_O_1
                    tagName = $"{xmlInfo.DB_Name}_O_1";
                    tag.SetAttribute("Name", tagName);
                    //设置连接点信息
                    tag.SetLinkValue("Connection", connection);
                    //tagTable的XML插入一个变量
                    tagTable.Insert(tag);

                    //递增
                    old = xmlInfo.DB_Name;
                    tagID++;
                    tagCount = 1;
                }

                //诊断接口超过8时
                if (tagCount > 8 && tagCount % 8 == 1)
                {
                    tag.SetId(tagID); //设置ID
                    //计算和设置地址%DB1.DBW1
                    string logicalAddress = xmlInfo.Offset.GetLogicalAddress(xmlInfo.DB_Number);
                    tag.SetAttribute("LogicalAddress", logicalAddress);
                    //计算和设置名称Pxxxx_O_1
                    tagName = $"{xmlInfo.DB_Name}_O_{(tagCount / 8) + (tagCount % 8)}";
                    tag.SetAttribute("Name", tagName);
                    //设置连接点信息
                    tag.SetLinkValue("Connection", connection);
                    //tagTable的XML插入一个变量
                    tagTable.Insert(tag);

                    //递增
                    tagID++;
                }
                
                if (old == xmlInfo.DB_Name)
                {
                    tagCount++;
                }

                //写入报警文本
                hmiAlarms_enUS.SetAlarmText(xmlInfo, "en-US", textID, tagName);
                hmiAlarms_zhCN.SetAlarmText(xmlInfo, "zh-CN", textID, tagName);

                textID++;
            }
        }
    }

    /// <summary>
    /// Xml文件信息
    /// </summary>
    public class XmlInfo
    {
        /// <summary>
        /// 诊断名称
        /// </summary>
        public string StateStruct                { get; set; }
        /// <summary>
        /// 诊断编号
        /// </summary>
        public string Number                     { get; set; }
        /// <summary>
        /// 诊断类型
        /// </summary>
        public string BlockTypeSupervisionNumber { get; set; }
        /// <summary>
        /// 报警文本-英文
        /// </summary>
        public string AlarmTextEn                { get; set; }
        /// <summary>
        /// 报警文本-中文
        /// </summary>
        public string AlarmTextZh                { get; set; }
        /// <summary>
        /// 偏移量
        /// </summary>
        public string Offset                     { get; set; }
        /// <summary>
        /// DB块名称
        /// </summary>
        public string DB_Name                    { get; set; }
        /// <summary>
        /// DB块编号
        /// </summary>
        public string DB_Number                  { get; set; }
    }
}