using MathCore.Hosting.WPF;

using TestWPF.ViewModels;

namespace TestWPF;

public class ServiceLocator : ServiceLocatorHosted
{
    public MainWindowViewModel MainModel => GetRequiredService<MainWindowViewModel>();
}