using System.Windows;
using System.Windows.Data;
using YukkuriMovieMaker.Commons;
using YukkuriMovieMaker.Controls;

namespace FontShuffle
{
    internal class FontShuffleControlAttribute : PropertyEditorAttribute2
    {
        public FontShuffleControlAttribute()
        {
            try
            {
                PropertyEditorSize = PropertyEditorSize.FullWidth;
                LogManager.WriteLog("FontShuffleControlAttributeが初期化されました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleControlAttribute初期化");
            }
        }

        public override FrameworkElement Create()
        {
            try
            {
                var control = new FontShuffleControl();
                LogManager.WriteLog("FontShuffleControlを作成しました");
                return control;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleControl作成");
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
                    LogManager.WriteLog("FontShuffleControlのバインディングを設定しました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleControlバインディング設定");
            }
        }

        public override void ClearBindings(FrameworkElement control)
        {
            try
            {
                if (control is FontShuffleControl selector)
                {
                    BindingOperations.ClearBinding(selector, FontShuffleControl.EffectProperty);
                    LogManager.WriteLog("FontShuffleControlのバインディングをクリアしました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleControlバインディングクリア");
            }
        }
    }
}