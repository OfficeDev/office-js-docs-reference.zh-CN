# <a name="metadata-element"></a>Metadata 元素

定义自定义函数在 Excel 中使用的元数据设置。

## <a name="attributes"></a>属性

无

## <a name="child-elements"></a>子元素

|  元素  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  是  | 使用自定义函数的 JSON 文件的资源 id 的字符串。 |

## <a name="example"></a>示例

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
