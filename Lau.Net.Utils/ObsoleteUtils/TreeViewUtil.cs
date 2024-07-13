//using System;
//using System.Collections.Generic;
//using System.Data;
//using System.Linq;
//using System.Text;
//using System.Windows.Forms;

//namespace Lau.Net.Utils
//{
//    public class TreeViewUtil
//    {
//        public virtual void AddSignleNode(TreeNode parentNode, string nodeName, string imageKey)
//        {
//            parentNode.Nodes.Add(nodeName).Name = nodeName;
//            parentNode.Nodes[nodeName].ImageKey = imageKey;
//        }

//        public virtual void AddNodes(TreeNode parentNode, DataTable dt, string imageKey)
//        {
//            foreach (DataRow dr in dt.Rows)
//            {
//                string nodeName = dr[0].ToString().Trim();
//                AddSignleNode(parentNode, nodeName, imageKey);
//            }
//        }

//        public virtual TreeNode AddSignleNode(TreeNode parentNode, string nodeText, string nodeName, string imageKey)
//        {
//            parentNode.Nodes.Add(nodeText).Name = nodeName;
//            parentNode.Nodes[nodeName].ImageKey = imageKey;
//            return parentNode.Nodes[nodeName];
//        }
//    }
//}
