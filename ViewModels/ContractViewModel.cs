using System.Collections.ObjectModel;
using ETA.Models;

namespace ETA.ViewModels;

public class ContractViewModel
{
    public ObservableCollection<Contract> Contracts { get; } = new();

    public ContractViewModel()
    {
        LoadContractData();
    }

    private void LoadContractData()
    {
        // ContractService.GetAllContracts() 등 기존 코드
    }
}