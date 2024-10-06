namespace Clippit.PowerPoint;

public class SlideSource
{
    public PmlDocument PmlDocument { get; set; }
    public int Start { get; set; }
    public int Count { get; set; }
    public bool KeepMaster { get; set; }

    public SlideSource(PmlDocument source, bool keepMaster)
    {
        PmlDocument = source;
        Start = 0;
        Count = int.MaxValue;
        KeepMaster = keepMaster;
    }

    public SlideSource(string fileName, bool keepMaster)
    {
        PmlDocument = new PmlDocument(fileName);
        Start = 0;
        Count = int.MaxValue;
        KeepMaster = keepMaster;
    }

    public SlideSource(PmlDocument source, int start, bool keepMaster)
    {
        PmlDocument = source;
        Start = start;
        Count = int.MaxValue;
        KeepMaster = keepMaster;
    }

    public SlideSource(string fileName, int start, bool keepMaster)
    {
        PmlDocument = new PmlDocument(fileName);
        Start = start;
        Count = int.MaxValue;
        KeepMaster = keepMaster;
    }

    public SlideSource(PmlDocument source, int start, int count, bool keepMaster)
    {
        PmlDocument = source;
        Start = start;
        Count = count;
        KeepMaster = keepMaster;
    }

    public SlideSource(string fileName, int start, int count, bool keepMaster)
    {
        PmlDocument = new PmlDocument(fileName);
        Start = start;
        Count = count;
        KeepMaster = keepMaster;
    }
}
