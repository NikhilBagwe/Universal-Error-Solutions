# Scenario :

## The ASP.NET application was working fine in local environment. But when hosted onto IIS of UAT it was throwing "Could not load file or assembly 'Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' or one of its dependencies" while downloading an existing Excel template file.

# Solution :

- The error was occuring because the above assembly file "version" was not correctly installed on UAT server.
- There was a mismatch between Local environment (ver 15.0.0) and UAT (ver 14.0.0) assembly file version.
- To solve this problem we added below piece of code in config file which basically told the application "at RUNTIME to use a different version of the assembly file if the specified version is not found".

```html
<configuration>
  <runtime>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Office.Interop.Excel" publicKeyToken="71e9bce111e9429c" culture="neutral" />
        <bindingRedirect oldVersion="0.0.0.0-15.0.0.0" newVersion="XX.XX.XX.XX" />
      </dependentAssembly>
    </assemblyBinding>
  </runtime>
</configuration>

```
- Replace XX.XX.XX.XX with the version number of the 'Microsoft.Office.Interop.Excel' assembly that is available on the server. This bindingRedirect element ensures that the application uses the available version if the requested version is not found.
- Thus, when the application didn't found the ver 15.0.0, at runtime it used the ver 14.0.0 which exsited on the UAT system and compiled the code.

## Conclusion : In version error, add code in config file to execute an alternative version of the required library/assembly file at RUNTIME.










