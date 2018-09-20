# <a name="scopes-element"></a>Scopes 元素

包含加载项需要拥有的对 Microsoft Graph 的访问权限。 Office 应用商店使用 Scopes 元素创建许可对话框。 当用户安装应用商店中的加载项时，系统会提示他们授予加载项对用户 Microsoft Graph 数据的指定访问权限。

## <a name="child-elements"></a>子元素

|  元素 |  类型  |  说明  |
|:-----|:-----|:-----|
|  **Scope**                |  string     |   Microsoft Graph 权限的名称，例如，Files.Read.All。 |

## <a name="example"></a>示例

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
