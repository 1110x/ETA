using System.Collections.ObjectModel;

namespace ETA.ViewModels
{
    public class TreeNode
    {
        public string Name { get; set; } = string.Empty;
        public bool IsExpanded { get; set; } = false;
        public ObservableCollection<TreeNode> Children { get; } = new ObservableCollection<TreeNode>();
    }
}