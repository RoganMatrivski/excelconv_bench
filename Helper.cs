// using SixLabors.Fonts;
// using SixLabors.ImageSharp;

// using System.Drawing;
// using SixLabors.ImageSharp.Drawing;

using SkiaSharp;

public static class Helper
{
    // public static Font GetSixLaborsFont(string family = "Carlito", int size = 11)
    // {
    //     var fontcoll = SystemFonts.Collection;
    //     var fontfamily = !string.IsNullOrWhiteSpace(family) ? fontcoll.Get(family) : SystemFonts.Collection.Families.FirstOrDefault();

    //     return new Font(fontfamily, size);
    // }

    // /**
    // * Excel measures columns in units of 1/256th of a character width
    // * but the docs say nothing about what particular character is used.
    // * '0' looks to be a good choice.
    // */
    private const char defaultChar = '0';

    public static SKSize MeasureString(string text, SKTypeface? typeface = null, float fontSize = 11)
    {
        var font = typeface ?? GetDefaultTypeface();

        using (var paint = new SKPaint
        {
            Typeface = font,
            TextSize = fontSize,
            IsAntialias = true,
            SubpixelText = true,
        })
        {
            var bounds = new SKRect();
            paint.MeasureText(text, ref bounds);

            return new SKSize(bounds.Width, bounds.Height);
        }
    }

    public static SKTypeface GetDefaultTypeface(string family = "Carlito") => SKFontManager.Default.MatchFamily(family);

    public static int GetDefaultCharWidth(SKTypeface? typeface = null, float fontSize = 11)
    {
        var width = MeasureString(new string(defaultChar, 1), typeface, fontSize).Width;
        return (int)Math.Round(width, 0, MidpointRounding.ToEven);
    }

    public static double GetStringWidth(string str, int defaultCharWidth, SKTypeface? typeface = null, float fontSize = 11)
    {
        // Split by newlines
        string[] lines = str.Split("\n".ToCharArray());

        double width = -1;
        foreach (var line in lines)
        {
            // Add some single char padding
            var txt = line + defaultChar;

            // Measure string, and then round to nearest whole number
            double w = Math.Round(
                MeasureString(txt, typeface, fontSize).Width,
                0,
                MidpointRounding.ToEven
            );

            // If current string is larger, store it. Else, ignore
            width = Math.Max(
                width,
                w / defaultCharWidth // Store the character size instead of the overall string size.
            );
        }

        return width * 256; // Multiply by 256 because Excel stores width in units of 1/256ths.
    }
}