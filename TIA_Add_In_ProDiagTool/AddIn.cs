using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.SW.Blocks;
using System;
using System.Linq;
using System.Windows.Forms;

namespace TIA_Add_In_ProDiagTool
{
    public class AddIn : ContextMenuAddIn
    {
        private static TiaPortal _tiaPortal;
        
        /// <summary>
        /// Base class for projects
        /// can be used in multi-user environment
        /// </summary>
        private ProjectBase _projectBase;

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
            assignment.Items.AddActionItem<IEngineeringObject>("如需配置，请选中实例数据块",
                menuSelectionProvider => { }, IsInstanceDB);
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
    }
}
