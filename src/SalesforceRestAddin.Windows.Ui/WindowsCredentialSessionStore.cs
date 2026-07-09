using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Windows.Ui;

/// <summary>
/// Stores refresh tokens in the Windows Credential Manager, scoped per login host.
/// </summary>
public sealed class WindowsCredentialSessionStore : ISessionCredentialStore
{
    private const string TargetPrefix = "SalesforceRestAddin/Session/";

    public static string TargetNameFor(string hostKey) => TargetPrefix + hostKey;

    public bool HasCredential(string hostKey) => Load(hostKey) is not null;

    public StoredSessionCredentials? Load(string hostKey)
    {
        if (string.IsNullOrWhiteSpace(hostKey))
        {
            return null;
        }

        if (!CredRead(TargetNameFor(hostKey), CredentialType.Generic, 0, out var credentialPtr))
        {
            return null;
        }

        try
        {
            var credential = Marshal.PtrToStructure<NativeCredential>(credentialPtr);
            if (credential.CredentialBlobSize == 0 || credential.CredentialBlob == IntPtr.Zero)
            {
                return null;
            }

            var jsonBytes = new byte[credential.CredentialBlobSize];
            Marshal.Copy(credential.CredentialBlob, jsonBytes, 0, jsonBytes.Length);
            var json = Encoding.UTF8.GetString(jsonBytes);
            return Deserialize(hostKey, json);
        }
        finally
        {
            CredFree(credentialPtr);
        }
    }

    public void Save(StoredSessionCredentials credentials)
    {
        if (credentials is null)
        {
            throw new ArgumentNullException(nameof(credentials));
        }

        var json = JsonSerializer.Serialize(new PersistedCredentialPayload
        {
            RefreshToken = credentials.RefreshToken,
            InstanceUrl = credentials.InstanceUrl,
            Id = credentials.Id,
            OAuthTokenHost = credentials.OAuthTokenHost,
        });

        var blob = Encoding.UTF8.GetBytes(json);
        var blobPointer = Marshal.AllocCoTaskMem(blob.Length);
        Marshal.Copy(blob, 0, blobPointer, blob.Length);

        var credential = new NativeCredential
        {
            AttributeCount = 0,
            Attributes = IntPtr.Zero,
            Comment = "SalesforceRestAddin Salesforce refresh token",
            TargetAlias = string.Empty,
            CredentialBlob = blobPointer,
            CredentialBlobSize = blob.Length,
            TargetName = TargetNameFor(credentials.HostKey),
            Type = CredentialType.Generic,
            Persist = CredentialPersistence.LocalMachine,
            UserName = credentials.HostKey,
        };

        try
        {
            if (!CredWrite(ref credential, 0))
            {
                throw new InvalidOperationException(
                    $"Failed to write Credential Manager entry for {credentials.HostKey} (Win32 error {Marshal.GetLastWin32Error()}).");
            }
        }
        finally
        {
            Marshal.FreeCoTaskMem(blobPointer);
        }
    }

    public void Delete(string hostKey)
    {
        if (string.IsNullOrWhiteSpace(hostKey))
        {
            return;
        }

        CredDelete(TargetNameFor(hostKey), CredentialType.Generic, 0);
    }

    private static StoredSessionCredentials? Deserialize(string hostKey, string json)
    {
        try
        {
            var payload = JsonSerializer.Deserialize<PersistedCredentialPayload>(json);
            var refreshToken = payload?.RefreshToken;
            if (payload is null || refreshToken is null || string.IsNullOrWhiteSpace(refreshToken))
            {
                return null;
            }

            return new StoredSessionCredentials
            {
                HostKey = hostKey,
                RefreshToken = refreshToken,
                OAuthTokenHost = string.IsNullOrWhiteSpace(payload.OAuthTokenHost) ? hostKey : payload.OAuthTokenHost,
                InstanceUrl = payload.InstanceUrl,
                Id = payload.Id,
            };
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private sealed class PersistedCredentialPayload
    {
        public string? RefreshToken { get; set; }

        public string? InstanceUrl { get; set; }

        public string? Id { get; set; }

        public string? OAuthTokenHost { get; set; }
    }

    private enum CredentialType : uint
    {
        Generic = 1,
    }

    private enum CredentialPersistence : uint
    {
        Session = 1,
        LocalMachine = 2,
        Enterprise = 3,
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NativeCredential
    {
        public uint Flags;
        public CredentialType Type;
        public string TargetName;
        public string Comment;
        public FileTime LastWritten;
        public int CredentialBlobSize;
        public IntPtr CredentialBlob;
        public CredentialPersistence Persist;
        public uint AttributeCount;
        public IntPtr Attributes;
        public string TargetAlias;
        public string UserName;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct FileTime
    {
        public uint LowDateTime;
        public uint HighDateTime;
    }

    [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredWrite([In] ref NativeCredential credential, uint flags);

    [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredRead(string target, CredentialType type, uint reservedFlag, out IntPtr credentialPtr);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool CredFree(IntPtr credentialPtr);

    [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredDelete(string target, CredentialType type, uint flags);
}
