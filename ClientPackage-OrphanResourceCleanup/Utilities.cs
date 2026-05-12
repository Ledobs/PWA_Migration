//Version: 16.0.7715.1200

using System;
using System.ServiceModel.Channels;
using System.Xml;

public static class MessageFaultHelper
{
    public static XmlElement GetMessageDetail(MessageFault fault)
    {
        return fault.GetDetail<XmlElement>();
    }
}