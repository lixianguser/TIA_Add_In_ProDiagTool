using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using System;
using System.Linq;
using System.Resources;
using System.Windows.Forms;

namespace TIA_Add_In_ProDiagTool
{
    public class AddIn : ContextMenuAddIn
    {
        private static TiaPortal _tiaPortal;
        
        /// <summary>
        /// Base class for projects
        /// can be use in multi-user environment
        /// </summary>
        private ProjectBase _projectBase;

        public AddIn(TiaPortal tiaPortal) : base("指定ProDiag FB")
        {
            _tiaPortal = tiaPortal;
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            addInRootSubmenu.Items.AddActionItem<InstanceDB>("指定ProDiag FB1", Assignment1_OnClick);
            addInRootSubmenu.Items.AddActionItem<InstanceDB>("指定ProDiag FB2", Assignment2_OnClick);
            addInRootSubmenu.Items.AddActionItem<InstanceDB>("指定ProDiag FB3", Assignment3_OnClick);
            addInRootSubmenu.Items.AddActionItem<InstanceDB>("无", AssignmentNone_OnClick);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("如需配置，请选中实例数据块",
                menuSelectionProvider => { }, IsInstanceDB);
        }

        //给实例DB块分配 ProDiag FB:
        //轮询所有blocks获取ProDiag FB语言的块
        //获取窗体选中的ProDiag FB块并指定到项目树中选中的实例块的ProDiag选项
        private void Assignment1_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            try
            { 
                // Multi-user support
                // If TIA Portal is in multi user environment (connected to project server)
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
                            //string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                            //指定AssignedProDiagFB
                            //if (assignedProDiagFB != "ProDiag_FB2")
                            //{
                            instanceDb.SetAttribute("AssignedProDiagFB", "ProDiag_FB1");
                            //}
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
                System.Windows.Forms.MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Assignment2_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider)
        {
            
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            try
            { 
                // Multi-user support
                // If TIA Portal is in multi user environment (connected to project server)
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
                            //string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                            //指定AssignedProDiagFB
                            //if (assignedProDiagFB != "ProDiag_FB2")
                            //{
                            instanceDb.SetAttribute("AssignedProDiagFB", "ProDiag_FB2");
                            //}
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
                System.Windows.Forms.MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Assignment3_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            try
            { 
                // Multi-user support
                // If TIA Portal is in multi user environment (connected to project server)
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
                            //string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                            //指定AssignedProDiagFB
                            //if (assignedProDiagFB != "ProDiag_FB2")
                            //{
                            instanceDb.SetAttribute("AssignedProDiagFB", "ProDiag_FB3");
                            //}
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
                System.Windows.Forms.MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AssignmentNone_OnClick(MenuSelectionProvider<InstanceDB> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            try
            {
                //轮询选中的实例块
                foreach (InstanceDB instanceDb in menuSelectionProvider.GetSelection())
                {
                    //string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                    //指定AssignedProDiagFB
                    //if (assignedProDiagFB != "ProDiag_FB2")
                    //{
                    instanceDb.SetAttribute("AssignedProDiagFB", "");
                    //}
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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
    }
}
