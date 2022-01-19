using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Windows.Markup;

using MathCore.WPF.ViewModels;

namespace TestWPF.ViewModels;

[MarkupExtensionReturnType(typeof(DataValueViewModel))]
public class DataValueViewModel : ViewModel
{
    #region Time : double - Параметр

    /// <summary>Параметр</summary>
    private double _Time;

    /// <summary>Параметр</summary>
    [Display(Name = "Время", Description = "Значение времени"), ReadOnly(true)]
    public double Time { get => _Time; set => Set(ref _Time, value); }

    #endregion

    #region Value : double - Значение

    /// <summary>Значение</summary>
    private double _Value;

    /// <summary>Значение</summary>
    [Display(Name = "Значение", Description = "Значение величины")]
    [DisplayFormat(DataFormatString = "0.00")]
    public double Value { get => _Value; set => Set(ref _Value, value); }

    #endregion

    #region Description : string - Описание

    /// <summary>Описание</summary>
    private string _Description;

    /// <summary>Описание</summary>
    [Display(Name = "Описание", Description = "Описание величины", AutoGenerateField = false)]
    public string Description { get => _Description; set => Set(ref _Description, value); }

    #endregion
}