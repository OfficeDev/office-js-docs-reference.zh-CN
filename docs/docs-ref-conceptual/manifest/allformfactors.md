# <a name="allformfactors-element"></a>AllFormFactors 元素

指定加载项的所有外观设置。 目前，仅使用**AllFormFactors**的功能是自定义函数。 使用自定义的函数时， **AllFormFactors**是必需的元素。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  是 |  定义加载项用于公开功能的位置。 |

## <a name="allformfactors-example"></a>AllFormFactors 示例

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
