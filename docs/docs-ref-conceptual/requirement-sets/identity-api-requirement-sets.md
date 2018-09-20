# <a name="identity-api-requirement-sets"></a>Identity API 要求集

要求集是指各组已命名的 API 成员。 Office 加载项使用清单中指定要求集，或使用运行时检查以确定是否的 Office 主机支持外接程序需要的 Api。 有关详细信息，请参阅[Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Office 加载项运行跨多个版本的 Office。 下表列出标识 API 要求集，Office 应用程序支持的要求集和生成或版本号码的 Office 主机应用程序。

|  要求集  | Office 2013 for Windows | 用于 Windows 的 office 365   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com & Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | 不适用 | 预览**和 #42;** | 即将推出 | 预览**和 #42;**| 可用 | 可用| 即将推出 | 即将推出 |

> **= #42;** 预览阶段中，Identity API 支持 Windows 2016 和 Mac 上只使用快速选项的内部程序中的用户。 若要加入内部程序，请参阅[为 Office 内幕](https://products.office.com/office-insider?tab=tab-1)。 要切换到快速跟踪，请参阅[Fast 内幕](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961)。

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Office Online Server 概述](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

有关公用 API 要求集的信息，请参阅 [Office 公用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="identityapi-11"></a>IdentityAPI 1.1 

单一登录 IdentityAPI 1.1 是 API 的第一个版本。 有关 API 的详细信息，请参阅`getAccessTokenAsync` [Office.Auth](/javascript/api/office/office.auth)参考主题中的方法。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
