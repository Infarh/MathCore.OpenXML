using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using System.Windows.Markup;

using MathCore.DI;
using MathCore.WPF.Commands;
using MathCore.WPF.ViewModels;

namespace TestWPF.ViewModels;

[Service]
[MarkupExtensionReturnType(typeof(MainWindowViewModel))]
public class MainWindowViewModel : TitledViewModel
{
    public MainWindowViewModel() => Title = "Главное окно";

    #region Visible : bool - Видимость окна

    /// <summary>Видимость окна</summary>
    private bool _Visible = true;

    /// <summary>Видимость окна</summary>
    public bool Visible { get => _Visible; set => Set(ref _Visible, value); }

    #endregion

    public ObservableCollection<DataValueViewModel> Values { get; } =
        Enumerable.Range(1, 100)
           .Select(i => new DataValueViewModel
            {
                Time = i,
                Value = i * 1000,
                Description = $"Описание {i}",
            })
           .ToObservableCollection();

    #region Command HideWindowCommand - Скрыть окно

    /// <summary>Скрыть окно</summary>
    private LambdaCommand _HideWindowCommand;

    /// <summary>Скрыть окно</summary>
    public ICommand HideWindowCommand => _HideWindowCommand
        ??= new(OnHideWindowCommandExecuted, CanHideWindowCommandExecute);

    /// <summary>Проверка возможности выполнения - Скрыть окно</summary>
    private bool CanHideWindowCommandExecute() => true;

    /// <summary>Логика выполнения - Скрыть окно</summary>
    private void OnHideWindowCommandExecuted()
    {
        Visible = false;
    }

    #endregion
}