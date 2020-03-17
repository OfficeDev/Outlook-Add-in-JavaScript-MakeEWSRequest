---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: O código JavaScript deste exemplo mostra uma solicitação simples para o assunto da mensagem de e-mail atual. Ele demonstra as etapas necessárias para criar uma solicitação do serviço Web do Exchange (EWS) e as práticas recomendadas para fazer a solicitação.
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:32:51 PM
---
# Suplemento do Outlook: Fazer uma solicitação de Serviço Web do Exchange pelo Outlook

**Índice**

* [Resumo](#summary)
* [Pré-requisitos](#prerequisites)
* [Componentes principais do exemplo](#components)
* [Descrição do código](#codedescription)
* [Criar e depurar](#build)
* [Solução de problemas](#troubleshooting)
* [Perguntas e comentários](#questions)
* [Recursos adicionais](#additional-resources)

<a name="summary"></a>
## Resumo
O código JavaScript deste exemplo mostra uma solicitação simples para o assunto da mensagem de e-mail atual. Ele demonstra as etapas necessárias para criar uma solicitação de EWS (Serviço Web do Exchange) e as práticas recomendadas para fazer a solicitação.

<a name="prerequisites"></a>
## Pré-requisitos ##

Esse exemplo requer o seguinte:  

  - Visual Studio 2013 com Atualização 5 ou Visual Studio 2015.  
  - Um computador executa o Exchange 2013 com pelo menos uma conta de e-mail ou uma conta do Office 365. Você pode [participar do Programa de Desenvolvedores do Office 365 e obter uma assinatura gratuita de um ano do Office 365](https://aka.ms/devprogramsignup).
  - Qualquer navegador que ofereça suporte ECMAScript 5.1, HTML5 e CSS3, como o Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6 ou uma versão posterior desses navegadores.
  - Familiaridade com a programação em JavaScript e serviços Web.

<a name="components"></a>
## Componentes principais do exemplo
A solução do exemplo contém os seguintes arquivos:

- MakeEwsRequestManifest.xml: O arquivo de manifesto do suplemento do Outlook.
- AppRead\Home\Home.html: A interface do usuário HTML do suplemento do Outlook.
- AppRead\Home\Home.js: O arquivo JavaScript que manipula a solicitação e o uso da solicitação de serviços Web do Exchange (EWS). 

<a name="codedescription"></a>
##Descrição do código

O código que cria a solicitação XML do EWS inclui dois métodos. O primeiro método, `getSoapEnvelope()`, delimita um envelope SOAP em torno de uma solicitação de serviço Web. Como o envelope SOAP é o padrão para todas as solicitações EWS, esse método pode ser reutilizado para delimitar qualquer solicitação do EWS.

O segundo método, `getSubjectRequest()`, retorna a solicitação EWS para obter o campo assunto de um item. O parâmetro ID é o identificador de item do Exchange para o item solicitado. Observe o seguinte sobre a solicitação:

- O elemento `ItemShape` é usado para restringir a resposta à forma base `IdOnly`. Isso limita a resposta apenas ao identificador de item do item e impede que sejam enviados dados excessivos para o servidor. 
- O elemento `AdditionalProperties` é usado para adicionar o campo Assunto à resposta. Ao usar a forma de base `IdOnly` e uma lista de propriedades adicionais, você pode limitar o tamanho da resposta do servidor a apenas os dados que seu suplemento exige. 

O método `sendRequest()` é chamado quando você clica no botão Fazer a solicitação EWS na IU do suplemento. Ele obtém o identificador do Exchange do item atual e o passa para os métodos `getSubjectRequest()` e `getSoapEnvelope()`, em seguida, faz uma chamada assíncrona para o servidor do Exchange usando o método `makeEwsRequestAsync`. O método `makeEwsRequestAsync` usa dois parâmetros: a solicitação EWS encapsulada em seu envelope SOAP e um método de retorno de chamada que é chamado quando a solicitação assíncrona para o EWS for concluída. Você pode adicionar um terceiro parâmetro opcional `userContext` ao método `makeEwsRequestAsync` caso tenha que fornecer informações adicionais para o método de retorno de chamada.

O método `callback()` é chamado com um parâmetro único, asyncResult. O objeto `asyncResult` tem dois membros:

- `value` - O conteúdo da resposta do EWS. 
- `context` - O parâmetro`userContext` passado para o método [makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2). 

O método `callback` no exemplo exibe o conteúdo da resposta em um elemento rolável `div` na interface do usuário, mas seu código pode usar a resposta de maneiras mais sofisticadas.

<a name="build"></a>
## Criar e depurar ##
**Observação**: O suplemento e-mail será ativado em qualquer mensagem de e-mail na caixa de entrada do usuário. Você pode facilitar o teste do suplemento enviando uma ou mais mensagens de e-mail para a sua conta de teste antes de executar o suplemento do exemplo.

1. Abra a solução no Visual Studio. Pressione F5 para criar e implementar o suplemento do exemplo.
2. Conecte-se a uma conta do Exchange fornecendo o endereço de e-mail e a senha de um servidor do Exchange 2013.
3. Permitir que o servidor configure a conta de e-mail.
4. No navegador, faça logon com a conta de e-mail digitando o nome e a senha da conta. 
5. Salvar uma mensagem na Caixa de entrada
6. Aguarde até que a barra de suplementos seja exibida sobre a mensagem.
7. Na barra de suplementos, clique em **MakeEWSRequest**.
8. Quando o suplemento de e-mail for exibido, clique no botão **Fazer solicitação EWS** para solicitar o assunto da mensagem atual do Exchange Server.
9. Examine o XML de resposta retornado pela solicitação.

<a name="troubleshooting"></a>
##Solução
de problemas Estes são erros comuns que podem ocorrer quando você usa o Outlook Web App para testar um suplemento de e-mail do Outlook:

- A barra de suplementos não será exibida quando uma mensagem for selecionada. Se isso ocorrer, reinicie o suplemento selecionando **Depurar - Parar a depuração** na janela do Visual Studio e, em seguida, pressione F5 para recriar e implementar o suplemento. 
- Pode ser que as alterações no código JavaScript não sejam selecionadas quando você implementar e executar o suplemento. Se as alterações não forem selecionadas, limpe o cache do navegador da Web selecionando **Ferramentas - Opções da Internet** e selecionando o botão **Excluir...**. Exclua os arquivos temporários da Internet e reinicie o suplemento. 

<a name="questions"></a>
##Perguntas e comentários##

- Se você tiver problemas para executar esse exemplo, [relate um problema](https://github.com/OfficeDev/Outlook-Add-in-Javascript-MakeEWSRequest/issues).
- Em geral, perguntas sobre o desenvolvimento de Suplementos do Office devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Não deixe de marcar as perguntas ou comentários com [office-addins].


<a name="additional-resources"></a>
## Recursos adicionais ##

- [Mais exemplos de Suplementos](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Explorar a API Gerenciada pelo EWS, EWS e serviços Web no Exchange](https://msdn.microsoft.com/library/office/jj536567(v=exchg.150).aspx)
- [método makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2)

## Direitos autorais
Copyright © 2015 Microsoft. Todos os direitos reservados.


Este projeto adotou o [Código de Conduta do Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
