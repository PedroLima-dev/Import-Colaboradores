# Lê credenciais do Windows Credential Manager e exporta como variáveis de ambiente
# Uso: powershell -File ler_credenciais.ps1

$code = @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class CredMan {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct CREDENTIAL {
        public int     Flags;
        public int     Type;
        public string  TargetName;
        public string  Comment;
        public long    LastWritten;
        public int     CredentialBlobSize;
        public IntPtr  CredentialBlob;
        public int     Persist;
        public int     AttributeCount;
        public IntPtr  Attributes;
        public string  TargetAlias;
        public string  UserName;
    }

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool CredRead(string target, int type, int flags, out IntPtr credential);

    [DllImport("advapi32.dll")]
    public static extern void CredFree(IntPtr credential);

    public static string[] GetCredential(string target) {
        IntPtr ptr = IntPtr.Zero;
        if (!CredRead(target, 1, 0, out ptr)) return null;
        var cred = (CREDENTIAL)Marshal.PtrToStructure(ptr, typeof(CREDENTIAL));
        byte[] blob = new byte[cred.CredentialBlobSize];
        Marshal.Copy(cred.CredentialBlob, blob, 0, cred.CredentialBlobSize);
        CredFree(ptr);
        return new string[] { cred.UserName, Encoding.Unicode.GetString(blob) };
    }
}
"@

Add-Type -TypeDefinition $code -Language CSharp

# Lê credencial AD
$ad = [CredMan]::GetCredential("UMUARAMA_AD_IMPORT")
if ($ad -eq $null) {
    Write-Host "ERRO_AD"
    exit 1
}

# Lê credencial do servidor
$sv = [CredMan]::GetCredential("10.56.43.28")
if ($sv -eq $null) {
    Write-Host "ERRO_SV"
    exit 1
}

# Exporta como variáveis de ambiente do processo pai (via arquivo temporário)
$tmpFile = $env:TEMP + "\umu_env.tmp"
[System.IO.File]::WriteAllLines($tmpFile, @(
    "UMU_USER=" + $ad[0],
    "UMU_PASS=" + $ad[1],
    "SERV_USER=" + $sv[0],
    "SERV_PASS=" + $sv[1],
    "AD_DISPLAY=" + $ad[0],
    "SV_DISPLAY=" + $sv[0]
), [System.Text.Encoding]::UTF8)

Write-Host "OK"
exit 0
