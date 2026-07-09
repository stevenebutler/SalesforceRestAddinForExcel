using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;
using SalesforceRestAddin.Core;

namespace SalesforceRestAddin;

/// <summary>
/// COMAddIns entry (<see cref="ComApiIdentity.ComAddInProgId"/>) — sets <c>Object</c> for VBA.
/// Ribbon UI is handled by <see cref="SalesforceRestAddinRibbon"/>.
/// </summary>
[ComVisible(true)]
[Guid(ComApiIdentity.ComAddInClassId)]
[ProgId(ComApiIdentity.ComAddInProgId)]
public sealed class SalesforceRestAddinComAddIn : ExcelComAddIn
{
    public override void OnConnection(
        object application,
        ext_ConnectMode connectMode,
        object addInInst,
        ref Array custom)
    {
        dynamic addIn = addInInst;
        addIn.Object = AddInHost.Api;
    }

    /// <summary>
    /// More reliable than <see cref="ExcelAddIn.AutoClose"/> on normal Excel quit — tear down
    /// WPF / COM so the process can exit (bugs.md #6).
    /// </summary>
    public override void OnBeginShutdown(ref Array custom)
    {
        AddInShutdown.Run();
        base.OnBeginShutdown(ref custom);
    }

    public override void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {
        AddInShutdown.Run();
        base.OnDisconnection(removeMode, ref custom);
    }
}
