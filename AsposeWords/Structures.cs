using OutSystems.ExternalLibraries.SDK;

namespace OSAsposeWords.Structures;


[OSStructure]
public struct AsposeResult
{
    public bool success;
    public string message;
    public byte[] document;



    public AsposeResult(bool _success, string _message, byte[] _doc)
    {
        this.success = _success;
        this.message = _message;
        this.document = _doc;

    }

}