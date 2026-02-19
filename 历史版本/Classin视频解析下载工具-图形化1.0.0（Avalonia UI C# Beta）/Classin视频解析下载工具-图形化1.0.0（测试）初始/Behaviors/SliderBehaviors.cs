using System;
using System.Windows.Input;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;

namespace VideoDownloader.Behaviors
{
    public static class SliderBehaviors
    {
        public static readonly AttachedProperty<ICommand?> SliderValueChangedCommandProperty =
            AvaloniaProperty.RegisterAttached<Slider, ICommand?>("SliderValueChangedCommand", typeof(SliderBehaviors));

        static SliderBehaviors()
        {
            SliderValueChangedCommandProperty.Changed.AddClassHandler<Slider>(OnSliderValueChangedCommandChanged);
        }

        public static ICommand? GetSliderValueChangedCommand(Slider element)
        {
            return element.GetValue(SliderValueChangedCommandProperty);
        }

        public static void SetSliderValueChangedCommand(Slider element, ICommand? value)
        {
            element.SetValue(SliderValueChangedCommandProperty, value);
        }

        private static void OnSliderValueChangedCommandChanged(Slider slider, AvaloniaPropertyChangedEventArgs e)
        {
            if (e.NewValue is ICommand command)
            {
                slider.ValueChanged += (s, args) => command.Execute(args.NewValue);
            }
        }
    }
}