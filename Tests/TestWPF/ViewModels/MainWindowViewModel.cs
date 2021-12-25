using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Markup;

using MathCore.DI;
using MathCore.WPF.ViewModels;

namespace TestWPF.ViewModels;

[Service]
[MarkupExtensionReturnType(typeof(MainWindowViewModel))]
public class MainWindowViewModel : TitledViewModel
{
    public MainWindowViewModel() => Title = "Главное окно";

    public ObservableCollection<DataValueViewModel> Values { get; } =
        Enumerable.Range(1, 100)
           .Select(i => new DataValueViewModel
            {
                Time = i,
                Value = i * 1000,
                Description = $"Описание {i}",
            })
           .ToObservableCollection();
}