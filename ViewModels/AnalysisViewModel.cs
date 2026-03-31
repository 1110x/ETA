using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;

namespace ETA.ViewModels
{
    public partial class AnalysisViewModel : ObservableObject
    {
        public ObservableCollection<AnalysisItem> AnalysisItems { get; } = [];
        public ObservableCollection<ContractPrice> Quotations { get; } = [];

        public ObservableCollection<string> Categories { get; } = [];

        [ObservableProperty]
        private AnalysisItem? selectedAnalysisItem;

        public AnalysisViewModel()
        {
            LoadAnalysisItems();
        }

        private void LoadAnalysisItems()
        {
            var items = AnalysisService.GetAllItems();
            foreach (var item in items)
            {
                AnalysisItems.Add(item);
                if (!string.IsNullOrEmpty(item.Category) && !Categories.Contains(item.Category))
                {
                    Categories.Add(item.Category);
                }
            }
            System.Diagnostics.Debug.WriteLine($"로드된 항목: {AnalysisItems.Count}, 카테고리: {Categories.Count}");
        }
    }
}