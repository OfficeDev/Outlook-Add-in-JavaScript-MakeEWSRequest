---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: Le code JavaScript de cet exemple décrit une demande simple pour l’objet de l’e-mail actuel. Il illustre les étapes nécessaires pour créer une demande de service web Exchange et les meilleures pratiques pour effectuer la demande.
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:32:51 PM
---
# Complément Outlook : Envoyer une demande de service web Exchange à partir d’Outlook

**Table des matières**

* [Résumé](#summary)
* [Conditions préalables](#prerequisites)
* [Composants clés de l’exemple](#components)
* [Description du code](#codedescription)
* [Création et débogage](#build)
* [Résolution des problèmes](#troubleshooting)
* [Questions et commentaires](#questions)
* [Ressources supplémentaires](#additional-resources)

<a name="summary"></a>
## Résumé
Le code JavaScript de cet exemple décrit une demande simple pour l’objet de l’e-mail actuel. Il illustre les étapes nécessaires pour créer une demande de service web Exchange et les meilleures pratiques pour effectuer la demande.

<a name="prerequisites"></a>
## Conditions préalables ##

Cet exemple nécessite les éléments suivants :  

  - Visual Studio 2013 avec la mise à jour 5 ou Visual Studio 2015.  
  - Un ordinateur exécutant Exchange 2013 avec au moins un compte de messagerie ou un compte Office 365. Vous pouvez [participer au programme pour les développeurs Office 365 et obtenir un abonnement gratuit d’un an à Office 365](https://aka.ms/devprogramsignup).
  - Tout navigateur qui prend en charge ECMAScript 5.1, HTML5 et CSS3, tel qu’Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6 ou une version ultérieure de ces navigateurs.
  - Être familiarisé avec les services web et de programmation JavaScript.

<a name="components"></a>
## Composants clés de l’exemple
La solution de l’exemple contient les fichiers suivants :

- MakeEwsRequestManifest.xml: Fichier manifeste pour le complément Outlook.
- AppRead\Home\Home.html : Interface utilisateur HTML pour le complément messagerie pour Outlook.
- AppRead\Home\Home.js : Le fichier JavaScript qui gère les demandes et l’utilisation de la demande Exchange Web Services (EWS). 

<a name="codedescription"></a>
##Description du code

Le code qui crée la requête XML EWS inclut deux méthodes. La première méthode, `getSoapEnvelope ()`, encapsule une enveloppe SOAP autour d’une demande de service Web. Étant donné que l’enveloppe SOAP est standard pour toutes les demandes EWS, cette méthode peut être réutilisée pour encapsuler une demande EWS.

La deuxième méthode, `getSubjectRequest ()`, renvoie la demande EWS pour obtenir le champ objet d’un élément. Le paramètre ID est l’identificateur d’élément Exchange pour l’élément demandé. Tenez compte des informations suivantes à propos de cette requête :

- L’élément `ItemShape` est utilisé pour limiter la réponse à la forme de base `IdOnly`. Cela limite la réponse uniquement à l’identificateur d’élément pour l’élément et empêche l’envoi de données excessives à partir du serveur. 
- L’élément `AdditionalProperties` est utilisé pour ajouter le champ Objet à la réponse. En utilisant la forme base de `IdOnly` et une liste de propriétés supplémentaires, vous pouvez limiter la taille de la réponse du serveur aux seules données requises par votre complément. 

La méthode `sendRequest ()` est appelée lorsque vous cliquez sur le bouton **faire une requête EWS** dans l’interface utilisateur du complément. Il récupère l’identificateur Exchange de l’élément actif et le transmet aux méthodes`getSubjectRequest ()` et `getSoapEnvelope ()`, puis effectue un appel asynchrone au serveur Exchange à l’aide de la méthode `makeEwsRequestAsync`. La méthode `makeEwsRequestAsync` prend deux paramètres : la demande EWS incluse dans son enveloppe SOAP, et une méthode de rappel appelée lorsque la requête asynchrone vers EWS est terminée. Vous pouvez ajouter un troisième paramètre `userContext` facultatif à la méthode `makeEwsRequestAsync` si vous devez fournir des informations supplémentaires à la méthode de rappel.

La méthode `callback ()` est appelée avec un seul paramètre, `asyncResult`. L’objet `asyncResult` a deux membres :

- `valeur` – contenu de la réponse de EWS. 
- `contexte` – le paramètre `userContext` transmis à la méthode [makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2). 

La méthode de `callback` dans l’exemple affiche le contenu de la réponse dans un élément de `div` de l’interface utilisateur à défilement variable, mais votre code peut utiliser la réponse de façon plus élaborée.

<a name="build"></a>
## Création et débogage ##
**Remarque** : Le complément messagerie sera activé sur tout courrier électronique figurant dans la boîte de réception de l’utilisateur. Vous pouvez simplifier le test du complément en envoyant un ou plusieurs courriers électroniques à votre compte de test avant d’exécuter l’exemple de complément.

1. Ouvrez la solution dans Visual Studio. Appuyez sur F5 pour créer et déployer l’exemple de complément.
2. Connectez-vous à un compte Exchange en fournissant l’adresse de courrier et le mot de passe d’un serveur Exchange 2013.
3. Autorisez le serveur à configurer le compte de messagerie.
4. Connectez-vous au compte de courrier en entrant le nom du compte et le mot de passe. 
5. Sélectionnez un message dans la boîte de réception.
6. Patientez jusqu’à ce que la barre du complément s’affiche au-dessus du message.
7. Dans la barre du complément, cliquez sur **MakeEWSRequest**.
8. Lorsque le complément courrier apparaît, cliquez sur le bouton **demander à EWS** pour demander l’objet du message actuel à partir du serveur Exchange Server.
9. Examinez le code XML de réponse renvoyé par la demande.

<a name="troubleshooting"></a>
##Dépannage
Voici quelques erreurs courantes qui peuvent se produire lors de l’utilisation d’Outlook Web App pour tester un complément messagerie pour Outlook :

- La barre du complément n’apparaît pas lorsqu’un message est sélectionné. Si c’est le cas, redémarrez l’application en sélectionnant **Debug – arrêter le débogage** dans la fenêtre Visual Studio, puis appuyez sur F5 pour regénérer et déployer le complément. 
- Les modifications apportées au code JavaScript peuvent ne pas être prises en compte lors du déploiement et de l’exécution du complément. Si les modifications ne sont pas prises en compte, effacez le cache du navigateur web en sélectionnant **Outils – Options Internet** puis en cliquant sur le bouton **Supprimer...** Supprimez les fichiers Internet temporaires, puis redémarrez le complément. 

<a name="questions"></a>
##Questions et commentaires##

- Si vous rencontrez des difficultés à l’exécution de cet exemple, veuillez [consigner un problème](https://github.com/OfficeDev/Outlook-Add-in-Javascript-MakeEWSRequest/issues).
- Si vous avez des questions générales sur le développement de compléments Office, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Posez vos questions ou envoyez vos commentaires en incluant la balise [office-addins].


<a name="additional-resources"></a>
## Ressources supplémentaires ##

- [Autres exemples de compléments](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Explorer l’API managée EWS, EWS et les services web dans Exchange](https://msdn.microsoft.com/library/office/jj536567(v=exchg.150).aspx)
- [Méthode makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2)

## Copyright
Copyright (c) 2015 Microsoft. Tous droits réservés.


Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
