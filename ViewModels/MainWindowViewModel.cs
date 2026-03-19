using System.Collections.ObjectModel;

namespace ETA.ViewModels
{
    public class TreeNode
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public ObservableCollection<TreeNode> Children { get; set; } = new();
    }

    public class ResultItem
    {
        public string Item { get; set; } = "";
        public double Value { get; set; }
    }

    public class MainWindowViewModel
    {
        public ObservableCollection<TreeNode> Data { get; set; }
        public ObservableCollection<ResultItem> Results { get; set; }

        public MainWindowViewModel()
        {
            Data = new ObservableCollection<TreeNode>();
            Results = new ObservableCollection<ResultItem>();

            LoadTree();
        }

        private void LoadTree()
        {
            var request = new TreeNode { Name = "의뢰 2026-001", Type = "의뢰" };

            var sampleA = new TreeNode { Name = "샘플 A", Type = "샘플" };
            sampleA.Children.Add(new TreeNode { Name = "COD", Type = "항목" });
            sampleA.Children.Add(new TreeNode { Name = "TN", Type = "항목" });

            var sampleB = new TreeNode { Name = "샘플 B", Type = "샘플" };
            sampleB.Children.Add(new TreeNode { Name = "PH", Type = "항목" });

            request.Children.Add(sampleA);
            request.Children.Add(sampleB);

            Data.Add(request);
        }

        public void LoadResult(string name)
        {
            Results.Clear();

            // 테스트 데이터 (나중에 DB로 교체)
            Results.Add(new ResultItem { Item = name, Value = 12.3 });
            Results.Add(new ResultItem { Item = name, Value = 7.8 });
        }
    }
}