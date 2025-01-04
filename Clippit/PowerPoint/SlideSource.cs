namespace Clippit.PowerPoint;

public class SlideSource(PmlDocument source, int start, int count, bool keepMaster)
{
    public PmlDocument PmlDocument { get; set; } = source;
    public int Start { get; set; } = start;
    public int Count { get; set; } = count;
    public bool KeepMaster { get; set; } = keepMaster;

    public SlideSource(PmlDocument source, bool keepMaster)
        : this(source, 0, int.MaxValue, keepMaster) { }

    public SlideSource(string fileName, bool keepMaster)
        : this(new PmlDocument(fileName), 0, int.MaxValue, keepMaster) { }

    public SlideSource(PmlDocument source, int start, bool keepMaster)
        : this(source, start, int.MaxValue, keepMaster) { }

    public SlideSource(string fileName, int start, bool keepMaster)
        : this(new PmlDocument(fileName), start, int.MaxValue, keepMaster) { }

    public SlideSource(string fileName, int start, int count, bool keepMaster)
        : this(new PmlDocument(fileName), start, count, keepMaster) { }
}
