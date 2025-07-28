using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

public static class LabelAnimator
{
    private static Dictionary<Label, AnimationData> animatedLabels = new Dictionary<Label, AnimationData>();
    private static Timer animationTimer;

    static LabelAnimator()
    {
        animationTimer = new Timer { Interval = 16 }; // ~60 FPS
        animationTimer.Tick += AnimationTimer_Tick;
        animationTimer.Start();
    }

    public static void SetupLabel(Label label,
                                Color normalColor,
                                Color hoverColor,
                                Color clickColor,
                                float originalSize = 9.75f,
                                float clickedSize = 8.75f)
    {
        if (animatedLabels.ContainsKey(label)) return;

        // Configuración inicial
        label.Font = new Font("Microsoft Sans Serif", originalSize, FontStyle.Bold);
        label.ForeColor = normalColor;
        label.Cursor = Cursors.Hand;

        // Registrar eventos
        label.MouseEnter += (s, e) => animatedLabels[label].IsHovering = true;
        label.MouseLeave += (s, e) => animatedLabels[label].IsHovering = false;
        label.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) animatedLabels[label].IsClicking = true; };
        label.MouseUp += (s, e) => { if (e.Button == MouseButtons.Left) animatedLabels[label].IsClicking = false; };

        // Agregar a la colección
        animatedLabels.Add(label, new AnimationData
        {
            NormalColor = normalColor,
            HoverColor = hoverColor,
            ClickColor = clickColor,
            OriginalFontSize = originalSize,
            ClickedFontSize = clickedSize,
            CurrentFontSize = originalSize
        });
    }

    private static void AnimationTimer_Tick(object sender, EventArgs e)
    {
        foreach (var entry in animatedLabels)
        {
            var label = entry.Key;
            var data = entry.Value;

            // Animación de color
            if (!data.IsClicking)
            {
                data.ColorTransitionProgress += data.IsHovering ? 0.1f : -0.1f;
                data.ColorTransitionProgress = Math.Max(0, Math.Min(1, data.ColorTransitionProgress));

                label.ForeColor = InterpolateColor(
                    data.NormalColor,
                    data.HoverColor,
                    data.ColorTransitionProgress);
            }
            else
            {
                label.ForeColor = data.ClickColor;
            }

            // Animación de tamaño
            data.CurrentFontSize += data.IsClicking ? -0.2f : 0.2f;
            data.CurrentFontSize = Math.Max(
                data.ClickedFontSize,
                Math.Min(data.OriginalFontSize, data.CurrentFontSize));

            label.Font = new Font(label.Font.FontFamily, data.CurrentFontSize, FontStyle.Bold);
        }
    }

    private static Color InterpolateColor(Color color1, Color color2, float progress)
    {
        return Color.FromArgb(
            (int)(color1.R + (color2.R - color1.R) * progress),
            (int)(color1.G + (color2.G - color1.G) * progress),
            (int)(color1.B + (color2.B - color1.B) * progress));
    }

    private class AnimationData
    {
        public Color NormalColor { get; set; }
        public Color HoverColor { get; set; }
        public Color ClickColor { get; set; }
        public float OriginalFontSize { get; set; }
        public float ClickedFontSize { get; set; }
        public float CurrentFontSize { get; set; }
        public float ColorTransitionProgress { get; set; }
        public bool IsHovering { get; set; }
        public bool IsClicking { get; set; }
    }
}