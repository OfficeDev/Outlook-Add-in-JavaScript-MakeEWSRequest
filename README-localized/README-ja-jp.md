---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: このサンプルの JavaScript コードは、現在の電子メール メッセージの件名に対するシンプルな要求を示しています。また、Exchange Web サービス (EWS) 要求の作成に必要な手順と、要求を行うためのベスト プラクティスも示しています。
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:32:51 PM
---
# Outlook アドイン:Outlook から Exchange Web サービス要求を行う

**目次**

* [概要](#summary)
* [前提条件](#prerequisites)
* [サンプルの主要なコンポーネント](#components)
* [コードの説明](#codedescription)
* [ビルドとデバッグ](#build)
* [トラブルシューティング](#troubleshooting)
* [質問とコメント](#questions)
* [その他のリソース](#additional-resources)

<a name="summary"></a>
## 概要
このサンプルの JavaScript コードは、現在の電子メール メッセージの件名に対するシンプルな要求を示しています。また、Exchange Web サービス (EWS) 要求の作成に必要な手順と、要求を行うためのベスト プラクティスも示しています。

<a name="prerequisites"></a>
## 前提条件 ##

このサンプルを実行するには次のものが必要です。  

  - Visual Studio 2013 更新プログラム 5 または Visual Studio 2015。  
  - 少なくとも 1 つのメール アカウントまたは Office 365 アカウントがある Exchange 2013 を実行するコンピューター。[Office 365 Developer プログラムに参加すると、Office 365 の 1 年間無料のサブスクリプションを取得](https://aka.ms/devprogramsignup)できます。
  - Internet Explorer 9、Chrome 13、Firefox 5、Safari 5.0.6、またはこれらのブラウザーの以降のバージョンなど、 ECMAScript 5.1、HTML5、および CSS3 をサポートする任意のブラウザー。
  - JavaScript プログラミングと Web サービスに関する知識。

<a name="components"></a>
## サンプルの主な構成要素
このサンプル ソリューションに含まれるファイルは次のとおりです。

- MakeEwsRequestManifest.xml:Outlook アドインのマニフェスト ファイル。
- AppRead\Home\Home.html:Outlook 用メール アドインの HTML ユーザー インターフェイス。
- AppRead\Home\Home.js:EWS 要求を処理および使用する JavaScript ファイル。 

<a name="codedescription"></a>
##コードの説明

EWS XML 要求を作成するコードには、2 のメソッドが含まれています。1 番目の `getSoapEnvelope()` は、Web サービス要求を SOAP エンベロープで囲むメソッドです。このメソッドを使用して、すべての EWS 要求の標準である SOAP エンベロープであらゆる EWS 要求を囲むことができます。

2 番目の `getSubjectRequest()` は、アイテムの件名フィールドを取得する EWS 要求を返すメソッドです。id パラメーターは、要求されたアイテムを示す Exchange アイテム識別子です。この要求については、次の点に注意してください。

- `ItemShape` 要素を使用して、BaseShape `IdOnly` に対する応答を制限します。これにより、応答がそのアイテムのアイテム識別子のみに制限され、サーバーから過剰なデータが送信されることはありません。 
- `AdditionalProperties` 要素を使用して、応答に Subject フィールドを追加します。BaseShape `IdOnly` やその他のプロパティの一覧を使用して、サーバーからの応答のサイズをアドインに必要なデータのみに制限できます。 

アドイン UI で [**Make EWS request (EWS 要求を行う)**] ボタンをクリックすると、`sendRequest()` メソッドが呼び出されます。現在のアイテムの Exchange 識別子を取得し、`getSubjectRequest()` と `getSoapEnvelope()` に渡します。次に `makeEwsRequestAsync` メソッドを使用して Exchange サーバーへの非同期呼び出しを行います。`makeEwsRequestAsync` メソッドは、2 つのパラメーターを受け取ります。それは、SOAP エンベロープに囲まれた EWS 要求と、EWS への非同期要求が完了したときに呼び出されるコールバック メソッドです。コールバック メソッドに追加情報が必要な場合は、3 番目のオプションの `userContext` パラメーターを `makeEwsRequestAsync` メソッドに追加できます。

`callback()` メソッドは、1 つのパラメーター `asyncResult` で呼び出されます。`asyncResult` オブジェクトには、2 つのメンバーが含まれています。

- `value` – EWS からの応答内容。 
- `context` – [makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2) メソッドに渡される `userContext` パラメーター。 

サンプルの `callback` メソッドは、UI でスクロール可能な `div` 要素に応答内容を表示していますが、実際のコードではより高度な方法で応答を使うことができます。

<a name="build"></a>
## ビルドとデバッグ ##
**注**: メール アドインは、ユーザーの受信トレイのすべてのメール メッセージで有効になります。サンプル アドインを実行する前に、1 つ以上のメール メッセージをテスト アカウントに送信しておくと、より簡単にアドインをテストできます。

1. Visual Studio でソリューションを開きます。F5 キーを押して、サンプル アドインをビルドおよび展開します。
2. Exchange 2013 Server 用のメール アドレスとパスワードを入力して Exchange アカウントに接続します。
3. サーバーがメール アカウントを構成できるようにします。
4. アカウント名とパスワードを入力して、メール アカウントにログオンします。 
5. 受信トレイのメッセージを選択します。
6. メッセージにアドイン バーが表示されるまで待ちます。
7. アドイン バーの **MakeEWSRequest** をクリックします。
8. メール アドインが表示されたら、[**Make EWS request (EWS 要求を行う)**] ボタンをクリックして、Exchange サーバーに現在のメッセージの件名を要求します。
9. 要求によって返された応答 XML を確認します。

<a name="troubleshooting"></a>
##トラブルシューティング
Outlook Web App を使用して Outlook のメール アドインをテストするときに発生する可能性がある一般的なエラーは次のとおりです。

- メッセージが選択されているときに、アドイン バーが表示されない。この問題が発生した場合は、Visual Studio ウィンドウで **[デバッグ]、[デバッグの停止]** の順に選択してアプリケーションを再起動し、次に F5 キーを押してアドインをリビルドして展開します。 
- アドインの展開と実行時に JavaScript コードの変更が認識されない場合がある。変更が認識されない場合は、**[ツール]、[インターネット オプション]** の順に選択し、[**削除…**] ボタンを選択して Web ブラウザーのキャッシュをクリアします。インターネット一時ファイルを削除してからアドインを再起動します。 

<a name="questions"></a>
##質問とコメント##

- このサンプルの実行について問題がある場合は、[問題をログに記録](https://github.com/OfficeDev/Outlook-Add-in-Javascript-MakeEWSRequest/issues)してください。
- Office アドイン開発全般の質問については、「[Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)」に投稿してください。質問やコメントには、必ず "office-addins" のタグを付けてください。


<a name="additional-resources"></a>
## その他の技術情報 ##

- [その他のアドイン サンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Exchange の EWS Managed API、EWS、および Web サービスについて学ぶ](https://msdn.microsoft.com/library/office/jj536567(v=exchg.150).aspx)
- [makeEwsRequestAsync メソッド](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2)

## 著作権
Copyright (c) 2015 Microsoft.All rights reserved.


このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
