using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

public class ComObj
{
    public static ComObj Create(string progId)
    {
        var type = Type.GetTypeFromProgID(progId);
        if (type == null)
        {
            throw new Exception("Servers need to be installed" + progId + "To use this feature");
        }
        return new ComObj(Activator.CreateInstance(type));
    }

    private object _val;
    public object Val
    {
        get { return _val; }
    }
    public ComObj(object comObject)
    {
        _val = comObject;
    }

    public ComObj Call(string mehtod, params object[] args)
    {
        if (_val == null)
            return null;
        var ret = _val.GetType().InvokeMember(mehtod, BindingFlags.InvokeMethod, null, _val, args);
        return new ComObj(ret);
    }
    public ComObj this[string property]
    {
        get
        {
            if (_val == null)
                return null;
            var ret = _val.GetType().InvokeMember(property, BindingFlags.GetProperty, null, _val, null);
            return new ComObj(ret);
        }
        set
        {
            if (_val != null)
                _val.GetType().InvokeMember(property, BindingFlags.SetProperty, null, _val, new object[] { value.Val });
        }
    }

    public static void ConvertSvgtoVsdx(string svgFn, string desVsdFn)
    {
        var app = ComObj.Create("Visio.Application"); // Required Visio to be installed locally
        try
        {
            // Vidio Required to generate the vsdx file for visio
            app["Visible"] = new ComObj(false);
            var docs = app["Documents"];
            short visOpenHidden = 64, visOpenRO = 2;
            var doc = docs.Call("OpenEx", svgFn, visOpenHidden + visOpenRO);
            doc.Call("SaveAs", desVsdFn);
            doc.Call("Close");

            var win = app["Window"];
        }
        finally
        {
            app.Call("Quit");
        }

    }
}

