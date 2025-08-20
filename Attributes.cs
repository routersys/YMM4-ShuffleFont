using System.Windows;
using System.Windows.Data;
using YukkuriMovieMaker.Commons;
using YukkuriMovieMaker.Controls;
using System;
using System.Threading.Tasks;

namespace FontShuffle
{
    internal class FontShuffleControlAttribute : PropertyEditorAttribute2
    {
        public FontShuffleControlAttribute()
        {
            try
            {
                PropertyEditorSize = PropertyEditorSize.FullWidth;
                _ = Task.Run(() => LogManager.WriteLog("FontShuffleControlAttributeが初期化されました"));
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "FontShuffleControlAttribute初期化"));
            }
        }

        public override FrameworkElement Create()
        {
            try
            {
                var control = new FontShuffleControl();
                _ = Task.Run(() => LogManager.WriteLog("FontShuffleControlを作成しました"));
                return control;
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "FontShuffleControl作成"));
                return new FontShuffleControl();
            }
        }

        public override void SetBindings(FrameworkElement control, ItemProperty[] itemProperties)
        {
            try
            {
                if (control is FontShuffleControl selector && itemProperties.Length > 0)
                {
                    selector.Effect = itemProperties[0].PropertyOwner as FontShuffleEffect;
                    _ = Task.Run(() => LogManager.WriteLog("FontShuffleControlのバインディングを設定しました"));
                }
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "FontShuffleControlバインディング設定"));
            }
        }

        public override void ClearBindings(FrameworkElement control)
        {
            try
            {
                if (control is FontShuffleControl selector)
                {
                    BindingOperations.ClearBinding(selector, FontShuffleControl.EffectProperty);
                    _ = Task.Run(() => LogManager.WriteLog("FontShuffleControlのバインディングをクリアしました"));
                }
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "FontShuffleControlバインディングクリア"));
            }
        }
    }
}