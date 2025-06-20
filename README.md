# PowerShellScripts

### OneDrive削除

1. **Windowsキー**+**R**
2. **「cmd」** と入力
3. **「Ctrl」キー＋「Shift」キー＋「Enter」キー** ※「管理者として実行」のショートカットキー。
4. 以下のコードをcmdの中に貼り付けて、「Enter」で実行します。

```cmd
@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/sit-oss/PowerShellScripts/refs/heads/main/odkiller.ps1'))"
```

