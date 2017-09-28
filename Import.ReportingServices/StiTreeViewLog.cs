using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;

namespace Import.Rdl
{
    public class StiTreeViewLog : IStiTreeLog
    {
        #region Fields
        private List<TreeNode> stackOfNodes = new List<TreeNode>();
        private TreeView treeView = null;
        #endregion

        #region Properties
        public TreeNode CurrentNode
        {
            get
            {
                return stackOfNodes[stackOfNodes.Count - 1];
            }
        }
        #endregion

        #region Methods
        public void OpenLog(string headerMessage)
        {
            TreeNode node = new TreeNode(headerMessage);
            treeView.Nodes.Add(node);
            stackOfNodes.Add(node);            
        }

        public void CloseLog()
        {
            treeView.ExpandAll();
            treeView.SelectedNode = this.CurrentNode;
            this.CurrentNode.EnsureVisible();
            stackOfNodes.Clear();
        }

        public void OpenNode(string message)
        {
            TreeNode node = new TreeNode(message);
            this.CurrentNode.Nodes.Add(node);
            stackOfNodes.Add(node);            
        }

        public void WriteNode(string message)
        {
            TreeNode node = new TreeNode(message);
            this.CurrentNode.Nodes.Add(node);
        }

        public void WriteNode(string message, object arg)
        {
            if (arg == null) WriteNode(message);
            else WriteNode(message + arg.ToString());
        }

        public void CloseNode()
        {
            stackOfNodes.RemoveAt(stackOfNodes.Count - 1);
        }
        #endregion

        public StiTreeViewLog(TreeView treeView)
        {
            this.treeView = treeView;
            this.treeView.Nodes.Clear();
        }
    }
}
