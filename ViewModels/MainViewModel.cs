using System.Collections.ObjectModel;
using ETA.Models;

namespace ETA.ViewModels;

public class MainViewModel
{
    public ObservableCollection<SampleData> Samples { get; set; }
        = new ObservableCollection<SampleData>();
}