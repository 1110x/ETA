using Avalonia.Controls;
using Avalonia.Controls.Models.TreeDataGrid;
using ETA.ViewModels;

namespace ETA.Views
{
    public partial class MainWindow : Window
    {
        private MainWindowViewModel _vm;

        public MainWindow()
        {
            InitializeComponent();

            _vm = new MainWindowViewModel();
            DataContext = _vm;

            // 🔥 TreeDataGrid 구성
            var source = new HierarchicalTreeDataGridSource<TreeNode>(_vm.Data)
            {
                Columns =
                {
                    new HierarchicalExpanderColumn<TreeNode>(
                        new TextColumn<TreeNode, string>("이름", x => x.Name),
                        x => x.Children
                    ),

                    new TextColumn<TreeNode, string>("유형", x => x.Type)
                }
            };

            treeGrid.Source = source;

            // 🔥 선택 이벤트
            treeGrid.RowSelection!.SelectionChanged += RowSelection_SelectionChanged;
        }

        private void RowSelection_SelectionChanged(object? sender, System.EventArgs e)
        {
            if (treeGrid.RowSelection?.SelectedItem is TreeNode node)
            {
                // 항목 클릭 시 결과 로드
                if (node.Type == "항목")
                {
                    _vm.LoadResult(node.Name);
                }
            }
        }
    }
}