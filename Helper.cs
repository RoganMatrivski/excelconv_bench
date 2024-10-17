// using SixLabors.Fonts;
// using SixLabors.ImageSharp;

// using System.Drawing;
// using SixLabors.ImageSharp.Drawing;

using System.Data;
using Humanizer;
using SkiaSharp;

public static class Helper
{
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

public static class IEnumerableExtensions
{
    public static IEnumerable<(int Index, T Item)> Indexed<T>(this IEnumerable<T> source)
    {
        return source.Select((item, index) => (index, item));
    }

    public static IEnumerable<DataColumn> AsEnumerable(this DataColumnCollection source)
    {
        return source.Cast<DataColumn>();
    }

    public static void PrintReadableTime(this TimeSpan time, string prepend = "", string append = "")
    {
        string readableTime = time.Humanize();
        Console.WriteLine($"{prepend}{readableTime}{append}");
    }

    // The fact that this is not built in is such a shame
    public static void ForEach<T>(this IEnumerable<T> @this, Action<T> action)
    {
        foreach (T item in @this)
        {
            action(item);
        }
    }
}