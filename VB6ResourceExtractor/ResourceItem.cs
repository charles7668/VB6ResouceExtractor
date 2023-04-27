namespace VB6ResourceExtractor;

public class ResourceItem
{
    public ResourceItem(string typeName, ResourceType resourceType, object content)
    {
        TypeName = typeName;
        ResourceType = resourceType;
        Content = content;
    }

    public object Content { get; set; }

    public ResourceType ResourceType { get; set; }

    public string TypeName { get; set; }
}