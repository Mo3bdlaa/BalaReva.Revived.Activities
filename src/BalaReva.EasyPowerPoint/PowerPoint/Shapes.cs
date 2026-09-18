using System.Collections.Generic;

namespace BalaReva.PowerPoint;

/// <summary>Font and emphasis of a text shape's text.</summary>
public sealed class ShapeFont
{
    /// <summary>Bold.</summary>
    public bool Bold { get; set; }

    /// <summary>Embossed.</summary>
    public bool Emboss { get; set; }

    /// <summary>Italic.</summary>
    public bool Italic { get; set; }

    /// <summary>Font name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Font size in points.</summary>
    public float Size { get; set; }

    /// <summary>Subscript.</summary>
    public bool Subscript { get; set; }

    /// <summary>Superscript.</summary>
    public bool Superscript { get; set; }

    /// <summary>Underlined.</summary>
    public bool Underline { get; set; }
}

/// <summary>One text shape on a slide: where it sits, what it says, how it looks.</summary>
public sealed class TextShape
{
    /// <summary>Height of the shape in points.</summary>
    public float BoundHeight { get; set; }

    /// <summary>Distance from the left edge of the slide, in points.</summary>
    public float BoundLeft { get; set; }

    /// <summary>Distance from the top edge of the slide, in points.</summary>
    public float BoundTop { get; set; }

    /// <summary>Width of the shape in points.</summary>
    public float BoundWidth { get; set; }

    /// <summary>The shape's text.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Font and emphasis of the text.</summary>
    public ShapeFont TextShapeFont { get; set; } = new();
}

/// <summary>The text shapes of one slide.</summary>
public sealed class SlideObject
{
    /// <summary>Every text shape on the slide, in the order PowerPoint reports them.</summary>
    public List<TextShape> TextShapes { get; set; } = [];
}
