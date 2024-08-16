using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.SW.Blocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using NPOI.SS.UserModel;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.HW;
using TIA_Add_In_Helper;

namespace TIA_Add_In_ProDiagTool
{
    public class AddIn : ContextMenuAddIn
    {
        private static TiaPortal         _tiaPortal;
        private        ProjectBase       _projectBase;
        private        string            _projectFolder;
        private        string            _exportFolder;
        private        List<XmlInfo>     _xmlInfos;

        public AddIn(TiaPortal tiaPortal) : base("ProDiag 工具")
        {
            _tiaPortal = tiaPortal;
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            Submenu assignment =addInRootSubmenu.Items.AddSubmenu("指定ProDiagFB");
            assignment.Items.AddActionItem<InstanceDB>("指定ProDiag FB1", Assignment1_OnClick);
            assignment.Items.AddActionItem<InstanceDB>("指定ProDiag FB2", Assignment2_OnClick);
            assignment.Items.AddActionItem<InstanceDB>("指定ProDiag FB3", Assignment3_OnClick);
            assignment.Items.AddActionItem<InstanceDB>("删除指定的ProDiag", AssignmentNone_OnClick);
            assignment.Items.AddActionItem<IEngineeringObject>("如需配置ProDiag，请选中实例数据块",
                menuSelectionProvider => { }, IsInstanceDB);
            
            Submenu proDiag4Hmi =addInRootSubmenu.Items.AddSubmenu("ProDiag4Hmi");
            proDiag4Hmi.Items.AddActionItem<DeviceItem>("生成数据", ProDiag4Hmi_OnClick);
            proDiag4Hmi.Items.AddActionItem<IEngineeringObject>("如需生成报警文本，请选中CPU设备项",
                menuSelectionProvider => { }, IsDeviceItem);
        }
        
        /// <summary>
        /// 给实例DB块指定ProDiag_FB1
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void Assignment1_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            try
            { 
                // Multi-user support
                // If TIA Portal is in multiuser environment (connected to project server)
                if (_tiaPortal.LocalSessions.Any())
                {
                    _projectBase = _tiaPortal.LocalSessions
                        .FirstOrDefault(s => s.Project != null && s.Project.IsPrimary)?.Project;
                }
                else
                {
                    // Get local project
                    _projectBase = _tiaPortal.Projects.FirstOrDefault(p => p.IsPrimary);
                }

                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("配置中……"))
                {
                    using (Transaction transaction = exclusiveAccess.Transaction(_projectBase,"指定FB监控为ProDiag_FB1"))
                    {
                        //轮询选中的实例块
                        foreach (InstanceDB instanceDb in menuSelectionProvider.GetSelection())
                        {
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                return;
                            }
                            string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                            //指定AssignedProDiagFB
                            if (assignedProDiagFB != "ProDiag_FB2")
                            {
                                exclusiveAccess.Text = instanceDb.Name + "指定FB监控为ProDiag_FB1";
                                instanceDb.SetAttribute("AssignedProDiagFB", "ProDiag_FB1");
                            }
                        }

                        if (transaction.CanCommit)
                        {
                            transaction.CommitOnDispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 给实例DB块指定ProDiag_FB2
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void Assignment2_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider)
        {
            
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            try
            { 
                // Multi-user support
                // If TIA Portal is in multiuser environment (connected to project server)
                if (_tiaPortal.LocalSessions.Any())
                {
                    _projectBase = _tiaPortal.LocalSessions
                        .FirstOrDefault(s => s.Project != null && s.Project.IsPrimary)?.Project;
                }
                else
                {
                    // Get local project
                    _projectBase = _tiaPortal.Projects.FirstOrDefault(p => p.IsPrimary);
                }

                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("配置中……"))
                {
                    using (Transaction transaction = exclusiveAccess.Transaction(_projectBase,"指定FB监控为ProDiag_FB2"))
                    {
                        //轮询选中的实例块
                        foreach (InstanceDB instanceDb in menuSelectionProvider.GetSelection())
                        {
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                return;
                            }
                            string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                            //指定AssignedProDiagFB
                            if (assignedProDiagFB != "ProDiag_FB2")
                            {
                                exclusiveAccess.Text = instanceDb.Name + "指定FB监控为ProDiag_FB2";
                                instanceDb.SetAttribute("AssignedProDiagFB", "ProDiag_FB2");
                            }
                        }

                        if (transaction.CanCommit)
                        {
                            transaction.CommitOnDispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 给实例DB块指定ProDiag_FB3
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void Assignment3_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            try
            { 
                // Multi-user support
                // If TIA Portal is in multiuser environment (connected to project server)
                if (_tiaPortal.LocalSessions.Any())
                {
                    _projectBase = _tiaPortal.LocalSessions
                        .FirstOrDefault(s => s.Project != null && s.Project.IsPrimary)?.Project;
                }
                else
                {
                    // Get local project
                    _projectBase = _tiaPortal.Projects.FirstOrDefault(p => p.IsPrimary);
                }

                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("配置中……"))
                {
                    using (Transaction transaction = exclusiveAccess.Transaction(_projectBase,"指定FB监控为ProDiag_FB3"))
                    {
                        //轮询选中的实例块
                        foreach (InstanceDB instanceDb in menuSelectionProvider.GetSelection())
                        {
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                return;
                            }
                            string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                            //指定AssignedProDiagFB
                            if (assignedProDiagFB != "ProDiag_FB2")
                            {
                                exclusiveAccess.Text = instanceDb.Name + "指定FB监控为ProDiag_FB3";
                                instanceDb.SetAttribute("AssignedProDiagFB", "ProDiag_FB3");
                            }
                        }

                        if (transaction.CanCommit)
                        {
                            transaction.CommitOnDispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 删除给实例DB块指定的ProDiag
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void AssignmentNone_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            try
            {
                // Multi-user support
                // If TIA Portal is in multiuser environment (connected to project server)
                if (_tiaPortal.LocalSessions.Any())
                {
                    _projectBase = _tiaPortal.LocalSessions
                        .FirstOrDefault(s => s.Project != null && s.Project.IsPrimary)?.Project;
                }
                else
                {
                    // Get local project
                    _projectBase = _tiaPortal.Projects.FirstOrDefault(p => p.IsPrimary);
                }

                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("配置中……"))
                {
                    using (Transaction transaction = exclusiveAccess.Transaction(_projectBase, "删除指定ProDiagFB"))
                    {
                        //轮询选中的实例块
                        foreach (InstanceDB instanceDb in menuSelectionProvider.GetSelection())
                        {
                            string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                            //指定AssignedProDiagFB
                            if (assignedProDiagFB != "")
                            {
                                exclusiveAccess.Text = instanceDb.Name + "删除指定FB监控";
                                instanceDb.SetAttribute("AssignedProDiagFB", "");
                            }
                        }
                        
                        if (transaction.CanCommit)
                        {
                            transaction.CommitOnDispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 生成触摸屏报警变量表和报警文本
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void ProDiag4Hmi_OnClick(MenuSelectionProvider<DeviceItem> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            DialogResult dialogResult =
                MessageBox.Show("请确认以下数据:\r\n"
                                + "1:PLC中包含ProDiagFB块，并确认相关实例DB已指定ProDiag。\r\n"
                                + "2:PLC已完全编译，并且无错误。\r\n"
                                + "3:检查'PLC监控和报警'中'监控定义'，所有的定义均包含对应中英文报警文本。\r\n"
                                + "按下'确认'键继续使用工具。", "提醒",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);

            if (dialogResult != DialogResult.OK)
            {
                return;
            }

            try
            {
                using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("初始化……"))
                {
                    //获取项目
                    _projectBase   = _tiaPortal.ProjectBase();
                    _projectFolder = _projectBase?.Path.Directory?.FullName;


                    //获取所有设备，并把名称和Device写入到数据
                    var deviceInfos = _projectBase.GetDeviceInfos();

                    exclusiveAccess.Text = "请选择目标触摸屏设备";

                    //新建触摸屏选择窗体
                    MainForm mainForm = new MainForm();
                    mainForm.SetDevices(deviceInfos); //填充数据
                    if (mainForm.ShowDialog() == DialogResult.OK)
                    {
                        // using (Transaction transaction =
                        //        exclusiveAccess.Transaction(_projectBase, "生成报警文本"))
                        // {
                        foreach (DeviceItem deviceItem in menuSelectionProvider.GetSelection())
                        {
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                break;
                            }

                            #region 处理PLC数据

                            exclusiveAccess.Text = "导出PLC数据";
                            //定义导出文件的保存文件夹路径
                            _exportFolder = $@"{_projectFolder}\ProDiag4Hmi";
                            //处理PLC数据
                            deviceItem.PlcDataAnalyze(_exportFolder, out _xmlInfos);

                            #endregion

                            #region 处理Hmi数据

                            exclusiveAccess.Text = "处理Hmi数据";

                            //处理触摸屏数据
                            HmiTarget hmiTarget = mainForm.device.GetHmiTarget();
                            hmiTarget.HmiDataAnalyze(_exportFolder,
                                _xmlInfos,
                                out XmlDocument tagTable,
                                out IWorkbook hmiAlarms_enUS,
                                out IWorkbook hmiAlarms_zhCN);

                            #endregion

                            //导入变量表
                            exclusiveAccess.Text = "导入Hmi变量表";
                            string exportPath = $@"{_projectFolder}\{tagTable.GetAttribute("Name")}.xml";
                            tagTable.Save(exportPath);
                            hmiTarget.TagFolder.ImportInfo(exportPath);

                            //保存文报警文本
                            exclusiveAccess.Text = "保存报警文本";
                            exportPath           = $@"{_projectFolder}\HMIAlarms-en_US.xlsx";
                            hmiAlarms_enUS.Save(exportPath);
                            exportPath = $@"{_projectFolder}\HMIAlarms-zh-CN.xlsx";
                            hmiAlarms_zhCN.Save(exportPath);

                            exclusiveAccess.Text = "正在完成";
                            MessageBox.Show($"报警文本.xlsx文件已保存到{_projectFolder}", "完成",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        //     if (transaction.CanCommit)
                        //     {
                        //         transaction.CommitOnDispose();
                        //     }
                        // }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
            finally
            {
                // 删除导出的文件夹
                Handle.DeleteDirectoryAndContents(_exportFolder);
            }
        }

        /// <summary>
        /// 判断当前选项是否为实例DB块
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns></returns>
        private static MenuStatus IsInstanceDB(MenuSelectionProvider menuSelectionProvider)
        {
            var show = false;

            foreach (IEngineeringObject engineeringObject in menuSelectionProvider.GetSelection())
            {
                if (!(engineeringObject is InstanceDB))
                {
                    show = true;
                    break;
                }
            }

            return show ? MenuStatus.Disabled : MenuStatus.Hidden;
        }
        
        /// <summary>
        /// 判断当前选项是否为设备项
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns></returns>
        private static MenuStatus IsDeviceItem(MenuSelectionProvider menuSelectionProvider)
        {
            var show = false;

            foreach (IEngineeringObject engineeringObject in menuSelectionProvider.GetSelection())
            {
                if (!(engineeringObject is DeviceItem))
                {
                    show = true;
                    break;
                }
            }

            return show ? MenuStatus.Disabled : MenuStatus.Hidden;
        }
    }
}
