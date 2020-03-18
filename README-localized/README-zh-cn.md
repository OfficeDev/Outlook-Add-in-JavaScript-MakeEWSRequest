---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: 本示例中的 JavaScript 代码显示针对当前电子邮件主题发出的一个简单请求。它演示创建 Exchange Web 服务 (EWS) 请求所需的步骤和发出请求的最佳做法。
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:32:51 PM
---
# Outlook 加载项：从 Outlook 发出 Exchange Web 服务请求

**目录**

* [摘要](#summary)
* [先决条件](#prerequisites)
* [示例主要组件](#components)
* [代码说明](#codedescription)
* [构建和调试](#build)
* [疑难解答](#troubleshooting)
* [问题和意见](#questions)
* [其他资源](#additional-resources)

<a name="summary"></a>
## 摘要
本示例中的 JavaScript 代码显示针对当前电子邮件主题发出的一个简单请求。它演示创建 Exchange Web 服务 (EWS) 请求所需的步骤和发出请求的最佳做法。

<a name="prerequisites"></a>
## 先决条件 ##

此示例要求如下：  

  - Visual Studio 2013 Update 5 或 Visual Studio 2015。  
  - 运行至少具有一个电子邮件帐户或 Office 365 帐户的 Exchange 2013 的计算机。你可以[参加 Office 365 开发人员计划并获取为期 1 年的免费 Office 365 订阅](https://aka.ms/devprogramsignup)。
  - 任何支持 ECMAScript 5.1、HTML5 和 CSS3 的浏览器，如 Internet Explorer 9、Chrome 13、Firefox 5、Safari 5.0.6 以及这些浏览器的更高版本。
  - 熟悉 JavaScript 编程和 Web 服务。

<a name="components"></a>
## 示例主要组件
本示例解决方案包含以下文件：

- MakeEwsRequestManifest.xml:Outlook 加载项的清单文件。
- AppRead\Home\Home.html：Outlook 邮件加载项的 HTML 用户界面。
- AppRead\Home\Home.js：处理请求和使用 EWS 请求的 JavaScript 文件。 

<a name="codedescription"></a>
##代码说明

创建 EWS XML 请求的代码包含两种方法。第一种方法是 `getSoapEnvelope()`，封装 Web 服务请求至 SOAP 信封。由于 SOAP 信封是所有 EWS 请求的标配，此方法可重新用于封装任何 EWS 请求。

第二种方法是 `getSubjectRequest()`，返回 EWS 请求以获得项目的主题字段。ID 参数是所请求项目的 Exchange 项目标识符。关于请求，请注意以下几点：

- `ItemShape` 元素用于限制对基本形状 `IdOnly` 的响应。这限制响应至项目的标识符，同时能够防止过度的数据从服务器发回。 
- `AdditionalProperties` 元素用于将“主题”字段添加到响应。通过使用 `IdOnly` 基本形状和附加属性列表，可限定来自服务器的响应大小为加载项请求的数据。 

单击加载项用户界面中的“**创建 EWS 请求**”时，调用 `sendRequest()` 方法。它获取当前项目的 Exchange 标识符并传递至 `getSubjectRequest()` 和 `getSoapEnvelope()` 方法，随后通过使用 `makeEwsRequestAsync` 方法对 Exchange 服务器进行异步调用。`makeEwsRequestAsync` 采用两个参数：封装在 SOAP 信封中的 EWS 请求，和调用方法（完成对 EWS 的异步调用时被调用）。如果为回调方法提供了附加信息，可添加可选的 `userContext` 参数至 `makeEwsRequestAsync` 方法。

`callback()` 方法使用单独参数 `asyncResult` 进行调用。`asyncResult` 对象有两个成员：

- `value` – EWS 响应的内容。 
- `context` – `userContext` 参数被传递至 [makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2) 方法。 

示例中的“`回调`”方法采用用户界面中的可滚动 `div` 元素显示响应的内容，但代码可采用更复杂的方式使用响应。

<a name="build"></a>
## 构建和调试 ##
**注意**：用户收件箱中的任何电子邮件均会激活邮件加载项。在运行示例加载项之前，可以向测试帐户发送一封或多封电子邮件，以便更轻松地测试该加载项。

1. 打开 Visual Studio 中的解决方案。按 F5 生成并部署示例加载项。
2. 通过为 Exchange 2013 服务器提供电子邮件地址和密码连接至 Exchange 帐户。
3. 允许服务器配置邮件帐户。
4. 通过输入帐户名称和密码登录电子邮件帐户。 
5. 选择收件箱中的一封邮件。
6. 等待加载项栏出现在邮件上方。
7. 在“加载项”栏中，单击“**MakeEWSRequest**”。
8. 如果显示邮件加载项，单击 **生成 EWS 请求** ，以从 Exchange 服务器请求当前邮件的主题。
9. 查看请求返回的响应 XML。

<a name="troubleshooting"></a>
##疑难解答
以下是当你使用 Outlook Web App 测试 Outlook 的邮件加载项时可能发生的常见错误：

- 选中邮件后，不会显示加载项栏。如果发生此情况，请通过在 Visual Studio 窗口中选择“**调试 - 停止调试**”来重启应用程序，然后按 F5 重建并部署加载项。 
- 部署和运行加载项时，可能不会记录对 JavaScript 代码的更改。如果更改未记录，请清除 Web 浏览器上的缓存，方法是选择“**工具 - Internet 选项**”并单击“**删除…**”按钮。删除临时 Internet 文件，然后重启加载项。 

<a name="questions"></a>
##问题和意见##

- 如果你在运行此示例时遇到任何问题，请[记录问题](https://github.com/OfficeDev/Outlook-Add-in-Javascript-MakeEWSRequest/issues)。
- 与 Office 外接程序开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)。确保你的问题或意见使用 [Office 外接程序] 进行了标记。


<a name="additional-resources"></a>
## 其他资源 ##

- [更多加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [在 Exchange 中浏览 EWS 托管 API、EWS 和 Web 服务](https://msdn.microsoft.com/library/office/jj536567(v=exchg.150).aspx)
- [makeEwsRequestAsync 方法](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2)

## 版权信息
版权所有 (c) 2015 Microsoft。保留所有权利。


此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
