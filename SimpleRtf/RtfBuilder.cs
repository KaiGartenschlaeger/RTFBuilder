using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace SimpleRtf
{
    /// <summary>
    /// Erzeugt Text im RichText-Format.
    /// </summary>
    public class RtfBuilder
    {
        #region Fields

        private Dictionary<int, int> colorNbs;
        private Dictionary<string, int> fontNbs;

        private StringBuilder sbColorTbl;
        private StringBuilder sbFontTbl;
        private StringBuilder sbText;

        private Color defaultColor;
        private Font defaultFont;

        #endregion

        #region Constructor

        /// <summary>
        /// Erzeugt Text im RichText-Format.
        /// </summary>
        public RtfBuilder()
        {
            colorNbs = new Dictionary<int, int>();
            fontNbs = new Dictionary<string, int>();

            sbColorTbl = new StringBuilder();
            sbFontTbl = new StringBuilder();
            sbText = new StringBuilder();

            defaultColor = SystemColors.WindowText;
            defaultFont = SystemFonts.DefaultFont;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Liefert die Standardschriftart, die verwendet wird falls keine andere explizit angegeben wird, oder legt diese fest.
        /// </summary>
        public Font DefaultFont
        {
            get
            {
                return defaultFont;
            }

            set
            {
                defaultFont = value;
            }
        }

        /// <summary>
        /// Liefert die Standardfarbe, die verwendet wird falls keine andere explizit angegeben wird, oder legt diese fest.
        /// </summary>
        public Color DefaultColor
        {
            get
            {
                return defaultColor;
            }

            set
            {
                defaultColor = value;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Fügt Text hinzu.
        /// </summary>
        /// <param name="text">Hinzuzufügender Text</param>
        public void Append(string text)
        {
            AppendIntern(text, RtfFormat.Normal, Color.Black, false, null, false);
        }

        /// <summary>
        /// Fügt Text hinzu.
        /// </summary>
        /// <param name="text">Hinzuzufügender Text</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        public void Append(string text, RtfFormat format)
        {
            AppendIntern(text, format, Color.Black, false, null, false);
        }

        /// <summary>
        /// Fügt Text hinzu.
        /// </summary>
        /// <param name="text">Hinzuzufügender Text</param>
        /// <param name="color">Zu verwendende Farbe</param>
        public void Append(string text, Color color)
        {
            AppendIntern(text, RtfFormat.Normal, color, true, null, false);
        }

        /// <summary>
        /// Fügt Text hinzu.
        /// </summary>
        /// <param name="text">Hinzuzufügender Text</param>
        /// <param name="font">Zu verwendende Schriftart</param>
        public void Append(string text, Font font)
        {
            AppendIntern(text, RtfFormat.Normal, Color.Empty, false, font, false);
        }

        /// <summary>
        /// Fügt Text hinzu.
        /// </summary>
        /// <param name="text">Hinzuzufügender Text</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        /// <param name="color">Zu verwendende Farbe</param>
        public void Append(string text, RtfFormat format, Color color)
        {
            AppendIntern(text, format, color, true, null, false);
        }

        /// <summary>
        /// Fügt Text hinzu.
        /// </summary>
        /// <param name="text">Hinzuzufügender Text</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        /// <param name="font">Zu verwendende Schriftart</param>
        public void Append(string text, RtfFormat format, Font font)
        {
            AppendIntern(text, format, Color.Empty, false, font, false);
        }

        /// <summary>
        /// Fügt Text hinzu.
        /// </summary>
        /// <param name="text">Hinzuzufügender Text</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        /// <param name="color">Zu verwendende Farbe</param>
        /// <param name="font">Zu verwendende Schriftart</param>
        public void Append(string text, RtfFormat format, Color color, Font font)
        {
            AppendIntern(text, format, color, true, font, false);
        }

        /// <summary>
        /// Fügt formatierten Text hinzu
        /// </summary>
        /// <param name="formattext">Hinzuzufügendes Textformat</param>
        /// <param name="parameter">Zu verwendende Formattext Parameter</param>
        public void AppendFormat(string formattext, params object[] parameter)
        {
            AppendIntern(string.Format(formattext, parameter), RtfFormat.Normal, Color.Black, false, null, false);
        }

        /// <summary>
        /// Fügt formatierten Text hinzu
        /// </summary>
        /// <param name="formattext">Hinzuzufügendes Textformat</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        /// <param name="parameter">Zu verwendende Formattext Parameter</param>
        public void AppendFormat(string formattext, RtfFormat format, params object[] parameter)
        {
            AppendIntern(string.Format(formattext, parameter), format, Color.Black, false, null, false);
        }

        /// <summary>
        /// Fügt formatierten Text hinzu
        /// </summary>
        /// <param name="formattext">Hinzuzufügendes Textformat</param>
        /// <param name="color">Zu verwendende Farbe</param>
        /// <param name="parameter">Zu verwendende Formattext Parameter</param>
        public void AppendFormat(string formattext, Color color, params object[] parameter)
        {
            AppendIntern(string.Format(formattext, parameter), RtfFormat.Normal, color, true, null, false);
        }

        /// <summary>
        /// Fügt formatierten Text hinzu
        /// </summary>
        /// <param name="formattext">Hinzuzufügendes Textformat</param>
        /// <param name="font">Zu verwendende Schriftart</param>
        /// <param name="parameter">Zu verwendende Formattext Parameter</param>
        public void AppendFormat(string formattext, Font font, params object[] parameter)
        {
            AppendIntern(string.Format(formattext, parameter), RtfFormat.Normal, Color.Empty, false, font, false);
        }

        /// <summary>
        /// Fügt formatierten Text hinzu
        /// </summary>
        /// <param name="formattext">Hinzuzufügendes Textformat</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        /// <param name="color">Zu verwendende Farbe</param>
        /// <param name="parameter">Zu verwendende Formattext Parameter</param>
        public void AppendFormat(string formattext, RtfFormat format, Color color, params object[] parameter)
        {
            AppendIntern(string.Format(formattext, parameter), format, color, true, null, false);
        }

        /// <summary>
        /// Fügt formatierten Text hinzu
        /// </summary>
        /// <param name="formattext">Hinzuzufügendes Textformat</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        /// <param name="font">Zu verwendende Schriftart</param>
        /// <param name="parameter">Zu verwendende Formattext Parameter</param>
        public void AppendFormat(string formattext, RtfFormat format, Font font, params object[] parameter)
        {
            AppendIntern(string.Format(formattext, parameter), format, Color.Empty, false, font, false);
        }

        /// <summary>
        /// Fügt formatierten Text hinzu
        /// </summary>
        /// <param name="formattext">Hinzuzufügendes Textformat</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        /// <param name="color">Zu verwendende Farbe</param>
        /// <param name="font">Zu verwendende Schriftart</param>
        /// <param name="parameter">Zu verwendende Formattext Parameter</param>
        public void AppendFormat(string formattext, RtfFormat format, Color color, Font font, params object[] parameter)
        {
            AppendIntern(string.Format(formattext, parameter), format, color, true, font, false);
        }

        /// <summary>
        /// Fügt Text mit einen Zeilenumbruch hinzu
        /// </summary>
        /// <param name="text">Hinzuzufügendes Text</param>
        public void AppendLine(string text)
        {
            AppendIntern(text, RtfFormat.Normal, Color.Black, false, null, true);
        }

        /// <summary>
        /// Fügt Text mit einen Zeilenumbruch hinzu
        /// </summary>
        /// <param name="text">Hinzuzufügendes Text</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        public void AppendLine(string text, RtfFormat format)
        {
            AppendIntern(text, format, Color.Black, false, null, true);
        }

        /// <summary>
        /// Fügt Text mit einen Zeilenumbruch hinzu
        /// </summary>
        /// <param name="text">Hinzuzufügendes Text</param>
        /// <param name="color">Zu verwendende Farbe</param>
        public void AppendLine(string text, Color color)
        {
            AppendIntern(text, RtfFormat.Normal, color, true, null, true);
        }

        /// <summary>
        /// Fügt Text mit einen Zeilenumbruch hinzu
        /// </summary>
        /// <param name="text">Hinzuzufügendes Text</param>
        /// <param name="font">Zu verwendende Schriftart</param>
        public void AppendLine(string text, Font font)
        {
            AppendIntern(text, RtfFormat.Normal, Color.Empty, false, font, true);
        }

        /// <summary>
        /// Fügt Text mit einen Zeilenumbruch hinzu
        /// </summary>
        /// <param name="text">Hinzuzufügendes Text</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        /// <param name="color">Zu verwendende Farbe</param>
        public void AppendLine(string text, RtfFormat format, Color color)
        {
            AppendIntern(text, format, color, true, null, true);
        }

        /// <summary>
        /// Fügt Text mit einen Zeilenumbruch hinzu
        /// </summary>
        /// <param name="text">Hinzuzufügendes Text</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        /// <param name="font">Zu verwendende Schriftart</param>
        public void AppendLine(string text, RtfFormat format, Font font)
        {
            AppendIntern(text, format, Color.Empty, false, font, true);
        }

        /// <summary>
        /// Fügt Text mit einen Zeilenumbruch hinzu
        /// </summary>
        /// <param name="text">Hinzuzufügendes Text</param>
        /// <param name="format">Zu verwendende Schriftart</param>
        /// <param name="color">Zu verwendende Farbe</param>
        /// <param name="font">Zu verwendende Schriftart</param>
        public void AppendLine(string text, RtfFormat format, Color color, Font font)
        {
            AppendIntern(text, format, color, true, font, true);
        }

        /// <summary>
        /// Löscht den hinzugefügten Text.
        /// </summary>
        public void Clear()
        {
            colorNbs.Clear();
            fontNbs.Clear();

            sbColorTbl.Clear();
            sbFontTbl.Clear();

            sbText.Clear();

            AddFont(SystemFonts.DefaultFont);
        }

        /// <summary>
        /// Liefert den erzeugten Text.
        /// </summary>
        public string GetRtf()
        {
            StringBuilder result = new StringBuilder();

            //header
            result.Append(@"{\rtf1\ansi\ansicpg1252\deff0\deflang1031");

            //font table
            result.Append(@"{\fonttbl");
            //defaultFont
            result.Append(FontToRtf(defaultFont, 0));
            //user fonts
            if (fontNbs.Count > 0)
            {
                result.Append(sbFontTbl.ToString());
            }
            result.Append("}");

            //color table
            result.Append(@"{\colortbl;");
            //default color
            result.Append(ColorToRtf(defaultColor));
            //user colors
            if (colorNbs.Count > 0)
            {
                result.Append(sbColorTbl.ToString());
            }
            result.Append("}");

            //text
            result.Append(sbText.ToString());

            //headerclose
            result.Append("}");

            return result.ToString();
        }

        /// <summary>
        /// Liefert den erzeugten Text.
        /// </summary>
        public override string ToString()
        {
            return GetRtf();
        }

        #endregion

        #region Intern Methods

        private string EscapeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            else
            {
                text = text.Replace(@"\", @"\\");
                text = text.Replace(@"{", @"\{");
                text = text.Replace(@"}", @"\}");

                return text;
            }
        }

        private int AddColor(Color color)
        {
            int argb = color.ToArgb();

            if (argb == defaultColor.ToArgb())
            {
                //Default Font
                return 1;
            }
            else
            {
                if (colorNbs.ContainsKey(argb))
                {
                    // Font exists already in Fonttable
                    return colorNbs[argb];
                }
                else
                {
                    //New Font
                    int nb = colorNbs.Count + 2;

                    colorNbs.Add(argb, nb);
                    sbColorTbl.AppendFormat(ColorToRtf(color));

                    return nb;
                }
            }
        }
        private int AddFont(Font font)
        {
            if (defaultFont.Name == font.Name)
            {
                //Default Font
                return 0;
            }
            else
            {
                if (fontNbs.ContainsKey(font.Name))
                {
                    //Font exists alread in font table
                    return fontNbs[font.Name];
                }
                else
                {
                    //new font
                    int fontNb = fontNbs.Count + 1;

                    sbFontTbl.Append(FontToRtf(font, fontNb));
                    fontNbs.Add(font.Name, fontNb);

                    return fontNb;
                }
            }
        }

        private string ColorToRtf(Color color)
        {
            return string.Format(@"\red{0}\green{1}\blue{2};",
                color.R, color.G, color.B);
        }
        private string FontToRtf(Font font, int nb)
        {
            return string.Format(@"{{\f{1}\fnil\fcharset0 {0}}}", font.Name, nb);
        }

        private void AppendIntern(string text, RtfFormat format, Color color, bool useColor, Font font, bool newLine)
        {
            //get color nb
            int colorNb = 1;
            if (useColor)
            {
                colorNb = AddColor(color);
            }

            //get font
            int fontNb = 0;
            int fontSize = (int)(defaultFont.SizeInPoints * 2);
            if (font != null)
            {
                fontNb = AddFont(font);
                fontSize = (int)(font.SizeInPoints * 2);
            }

            //set formats
            if ((format & RtfFormat.Bold) == RtfFormat.Bold)
            {
                sbText.Append(@"\b");
            }
            if ((format & RtfFormat.Italic) == RtfFormat.Italic)
            {
                sbText.Append(@"\i");
            }
            if ((format & RtfFormat.Underline) == RtfFormat.Underline)
            {
                sbText.Append(@"\ul");
            }
            if ((format & RtfFormat.Strikeout) == RtfFormat.Strikeout)
            {
                sbText.Append(@"\strike");
            }

            //set color
            sbText.AppendFormat(@"\cf{0}", colorNb);

            //set font
            sbText.AppendFormat(@"\f{0}", fontNb);
            sbText.AppendFormat(@"\fs{0} ", fontSize);

            //add text
            sbText.Append(EscapeText(text));

            //restore to default color
            sbText.Append(@"\cf1");

            //restore to default font
            sbText.Append(@"\f0");
            sbText.AppendFormat(@"\fs{0} ", (int)(defaultFont.SizeInPoints * 2));

            //remove formats
            if ((format & RtfFormat.Bold) == RtfFormat.Bold)
            {
                sbText.Append(@"\b0");
            }
            if ((format & RtfFormat.Italic) == RtfFormat.Italic)
            {
                sbText.Append(@"\i0");
            }
            if ((format & RtfFormat.Underline) == RtfFormat.Underline)
            {
                sbText.Append(@"\ul0");
            }
            if ((format & RtfFormat.Strikeout) == RtfFormat.Strikeout)
            {
                sbText.Append(@"\strike0");
            }

            //insert new line
            if (newLine)
            {
                sbText.Append(@"\par ");
            }
        }

        #endregion
    }
}