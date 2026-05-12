//#Version: 16.0.7715.1200

using System;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Dispatcher;
using System.ServiceModel.Description;

public class UserDataRequestBehavior: IClientMessageInspector, IEndpointBehavior
{
    public string RequestDigest { get; set; }
    public string AuthCookie { get; set; }
    public Uri RequestUri { get; set; }

    #region IClientMessageInspector
    public void AfterReceiveReply(ref System.ServiceModel.Channels.Message reply, object correlationState)
    {
    }

    public object BeforeSendRequest(ref System.ServiceModel.Channels.Message request, IClientChannel channel)
    {
        if (!string.IsNullOrEmpty(RequestDigest) || !string.IsNullOrEmpty(AuthCookie))
        {
            HttpRequestMessageProperty property = new HttpRequestMessageProperty();

            if (!string.IsNullOrEmpty(RequestDigest))
            {
                property.Headers.Add("X-RequestDigest", RequestDigest);
            }
            if (!string.IsNullOrEmpty(AuthCookie))
            {
                CookieContainer cookies = new CookieContainer();
                cookies.SetCookies(RequestUri, AuthCookie);
                property.Headers.Add(HttpRequestHeader.Cookie, cookies.GetCookieHeader(RequestUri));
            }

            request.Properties.Add(HttpRequestMessageProperty.Name, property);
        }

        return null;
    }
    #endregion

    #region IEndpointBehavior
    public void AddBindingParameters(ServiceEndpoint endpoint, System.ServiceModel.Channels.BindingParameterCollection bindingParameters)
    {
    }

    public void ApplyClientBehavior(ServiceEndpoint endpoint, ClientRuntime clientRuntime)
    {
        clientRuntime.MessageInspectors.Add(this);
    }

    public void ApplyDispatchBehavior(ServiceEndpoint endpoint, EndpointDispatcher endpointDispatcher)  
    {
    }

    public void Validate(ServiceEndpoint endpoint)
    {
    }
    #endregion
}