using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using System;
using System.Windows.Forms;

namespace TIA_Add_In_ProDiagTool
{
    public class AddIn : ContextMenuAddIn
    {
        private static TiaPortal _tiaPortal;

        public AddIn(TiaPortal tiaPortal) : base("ProDiagTool")
        {
            _tiaPortal = tiaPortal;
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            addInRootSubmenu.Items.AddActionItem<InstanceDB>("指定ProDiag FB1", Assignment1_OnClick);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("指定ProDiag FB1",
                menuSelectionProvider => { }, IsInstanceDB);
            addInRootSubmenu.Items.AddActionItem<InstanceDB>("指定ProDiag FB2", Assignment2_OnClick);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("指定ProDiag FB2",
                menuSelectionProvider => { }, IsInstanceDB);
            addInRootSubmenu.Items.AddActionItem<InstanceDB>("指定ProDiag FB3", Assignment3_OnClick);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("指定ProDiag FB3",
                menuSelectionProvider => { }, IsInstanceDB);
            addInRootSubmenu.Items.AddActionItem<InstanceDB>("无", AssignmentNone_OnClick);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("无",
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
                //轮询选中的实例块
                foreach (InstanceDB instanceDb in menuSelectionProvider.GetSelection())
                {
                    //string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                    //指定AssignedProDiagFB
                    //if (assignedProDiagFB != "ProDiag_FB2")
                    //{
                        instanceDb.SetAttribute("AssignedProDiagFB", "ProDiag_FB1");
                    //}
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
                //轮询选中的实例块
                foreach (InstanceDB instanceDb in menuSelectionProvider.GetSelection())
                {
                    //string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                    //指定AssignedProDiagFB
                    //if (assignedProDiagFB != "ProDiag_FB2")
                    //{
                    instanceDb.SetAttribute("AssignedProDiagFB", "ProDiag_FB2");
                    //}
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
                //轮询选中的实例块
                foreach (InstanceDB instanceDb in menuSelectionProvider.GetSelection())
                {
                    //string assignedProDiagFB = instanceDb.GetAttribute("AssignedProDiagFB").ToString();
                    //指定AssignedProDiagFB
                    //if (assignedProDiagFB != "ProDiag_FB2")
                    //{
                    instanceDb.SetAttribute("AssignedProDiagFB", "ProDiag_FB3");
                    //}
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
