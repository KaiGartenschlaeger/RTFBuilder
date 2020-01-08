using System;

namespace SimpleRtf
{
    /// <summary>
    /// Legt das zu verwendende Textformat fest.
    /// </summary>
    [Flags]
    public enum RtfFormat
    {
        /// <summary>
        /// Normal
        /// </summary>
        Normal = 0,

        /// <summary>
        /// Fett
        /// </summary>
        Bold = 1,

        /// <summary>
        /// Kursiv
        /// </summary>
        Italic = 2,

        /// <summary>
        /// Unterstrichen
        /// </summary>
        Underline = 4,

        /// <summary>
        /// Durchgestrichen
        /// </summary>
        Strikeout = 8
    }
}