using System.Collections.ObjectModel;
using ETA.Models;

namespace ETA.ViewModels;

public class QuotationViewModel
{
    public ObservableCollection<ContractPrice> QuotationItems { get; } = new();

    public QuotationViewModel()
    {
        LoadQuotationData();
    }

    private void LoadQuotationData()
    {
        // ContractPriceService.GetAllContractPrices() 등 기존 코드
    }
}